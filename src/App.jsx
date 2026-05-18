import React, { useState, useRef, useCallback, useEffect, useMemo } from "react";
import ExcelJS from "exceljs";

// ─── Design tokens ────────────────────────────────────────────────────────────
const C = {
  bg: "#0f0f0f", surface: "#1a1a1a", surface2: "#141414",
  border: "#2a2a2a", orange: "#ff6b1a",
  green: "#2ecc71", red: "#e74c3c", yellow: "#f39c12", blue: "#3498db",
  text: "#e8e8e8", muted: "#666",
};

const S = {
  app: { fontFamily: "'Courier New', monospace", background: C.bg, minHeight: "100vh", color: C.text, maxWidth: 480, margin: "0 auto" },
  header: { background: C.surface, borderBottom: `2px solid ${C.orange}`, padding: "13px 16px", display: "flex", alignItems: "center", gap: 12, position: "sticky", top: 0, zIndex: 20 },
  logo: { width: 30, height: 30, background: C.orange, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, fontWeight: "bold", color: "#000", flexShrink: 0 },
  body: { padding: "14px 14px 90px" },
  card: { background: C.surface, border: `1px solid ${C.border}`, borderRadius: 2, marginBottom: 12, overflow: "hidden" },
  cardHead: { padding: "7px 13px", background: C.surface2, borderBottom: `1px solid ${C.border}`, fontSize: 9, letterSpacing: 2, color: C.orange, textTransform: "uppercase" },
  cardBody: { padding: 13 },
  label: { fontSize: 9, color: C.muted, letterSpacing: 1.5, textTransform: "uppercase", marginBottom: 4, display: "block" },
  input: (accent) => ({ width: "100%", background: "#111", border: `1px solid ${accent || C.border}`, borderRadius: 2, color: C.text, fontFamily: "inherit", fontSize: 13, padding: "8px 10px", outline: "none", boxSizing: "border-box" }),
  row: { display: "flex", gap: 9, marginBottom: 10 },
  col: { flex: 1, minWidth: 0 },
  btn: (v = "primary", s = "md") => ({
    background: v === "primary" ? C.orange : v === "danger" ? C.red : "transparent",
    border: v === "ghost" ? `1px solid ${C.border}` : "none",
    color: v === "primary" ? "#000" : C.text,
    fontFamily: "inherit", fontSize: s === "sm" ? 9 : 11, fontWeight: "bold",
    letterSpacing: 1.5, textTransform: "uppercase",
    padding: s === "sm" ? "6px 11px" : "10px 16px",
    borderRadius: 2, cursor: "pointer", width: "100%",
  }),
  imgPlaceholder: { width: "100%", aspectRatio: "4/3", background: "#111", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", cursor: "pointer", borderBottom: `1px solid ${C.border}` },
  imgPreview: { width: "100%", aspectRatio: "4/3", objectFit: "cover", display: "block", borderBottom: `1px solid ${C.border}` },
  tag: (color) => ({ display: "inline-block", padding: "2px 7px", background: color + "22", color, border: `1px solid ${color}44`, borderRadius: 2, fontSize: 9, fontWeight: "bold", letterSpacing: 1 }),
  th: { textAlign: "left", padding: "6px 8px", borderBottom: `1px solid ${C.border}`, color: C.orange, fontSize: 9, letterSpacing: 1.5, textTransform: "uppercase" },
  td: { padding: "6px 8px", borderBottom: `1px solid ${C.border}`, verticalAlign: "middle" },
};

const isEAN = (v) => /^\d{8,14}$/.test(v.trim());
const fmtPrice = (n) => Number(n).toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// ─── Anagrafica parser (ExcelJS — no vulnerabilities) ────────────────────────
function parseAnagrafica(file) {
  return new Promise((resolve, reject) => {
    const isCSV = file.name.toLowerCase().endsWith(".csv");
    const reader = new FileReader();

    reader.onload = async (e) => {
      try {
        const workbook = new ExcelJS.Workbook();

        if (isCSV) {
          // Parse CSV manually to avoid ExcelJS stream dependency in browser
          const text = new TextDecoder("utf-8").decode(e.target.result);
          const lines = text.split(/\r?\n/).filter(l => l.trim());
          if (lines.length < 2) { reject(new Error("File CSV vuoto")); return; }
          const sep = lines[0].includes(";") ? ";" : ",";
          const parseCSVLine = (l) => l.split(sep).map(c => c.replace(/^"|"$/g,"").trim());
          const hdrs = parseCSVLine(lines[0]);
          const ws = workbook.addWorksheet("CSV");
          ws.addRow(hdrs);
          for (let i = 1; i < lines.length; i++) ws.addRow(parseCSVLine(lines[i]));
        } else {
          await workbook.xlsx.load(e.target.result);
        }

        const ws = workbook.worksheets[0];
        if (!ws || ws.rowCount < 2) { reject(new Error("File vuoto o senza dati")); return; }

        // Build header index from row 1
        const hdrRow = ws.getRow(1);
        const colIndex = {}; // colName -> colNumber
        hdrRow.eachCell({ includeEmpty: true }, (cell, col) => {
          const key = String(cell.value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
          colIndex[key] = col;
        });

        const findCol = (...kw) => {
          for (const key of Object.keys(colIndex)) {
            if (kw.some(k => key.includes(k))) return colIndex[key];
          }
          return null;
        };

        const colSKU   = findCol("sku","codart","codicearticolo","codice","cod");
        const colEAN   = findCol("ean","barcode","codiceabarre","gtin","barre");
        const colDesc  = findCol("descri","denominaz","nome","articolo");
        const colPrice = findCol("prezzoac","costoac","acquisto","costo","prezzo","listino","price");

        if (!colSKU && !colEAN) { reject(new Error("Nessuna colonna SKU o EAN trovata nel file")); return; }

        const map = {};
        let count = 0;

        ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
          if (rowNum === 1) return;
          const cell = (col) => col ? String(row.getCell(col).value ?? "").trim() : "";
          const sku   = colSKU  ? cell(colSKU).toUpperCase() : "";
          const ean   = colEAN  ? cell(colEAN).replace(/\D/g, "") : "";
          const desc  = colDesc ? cell(colDesc) : "";
          const rawP  = colPrice ? cell(colPrice).replace(",", ".") : "";
          const p     = parseFloat(rawP);
          const price = colPrice ? (isNaN(p) ? null : p) : null;
          const entry = { sku, ean, desc, price };
          if (sku) map[sku] = entry;
          if (ean) map[ean] = entry;
          count++;
        });

        resolve({ map, count });
      } catch (err) { reject(err); }
    };

    reader.onerror = () => reject(new Error("Lettura file fallita"));
    // ArrayBuffer for xlsx, ArrayBuffer for CSV too (decoded inside)
    reader.readAsArrayBuffer(file);
  });
}

// ─── Claude Vision ────────────────────────────────────────────────────────────
async function countPiecesWithAI(base64, mediaType, identifier) {
  const resp = await fetch("/api/chat", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6", max_tokens: 200,
      system: `Sei un sistema di conteggio inventario per magazzino ricambi bici.
Conta i pezzi visibili nella foto del cassettone/ripiano.
Rispondi SOLO in JSON senza markdown:
{"count":<intero>,"confidence":"alta"|"media"|"bassa","note":"<max 60 char>"}
Se non riesci: count:-1.`,
      messages: [{ role: "user", content: [
        { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
        { type: "text", text: `Articolo: ${identifier || "n/d"}. Conta tutti i pezzi visibili.` }
      ]}],
    }),
  });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}: ${resp.statusText}`); // BUG FIX B
  const d = await resp.json();
  const raw = d.content?.find(b => b.type === "text")?.text || "{}";
const match = raw.match(/\{[\s\S]*\}/);
  if (!match) return { count: -1, confidence: "bassa", note: "Risposta AI non valida" };
  return JSON.parse(match[0]);
}

// ─── EAN web lookup ───────────────────────────────────────────────────────────
async function lookupEANweb(ean) {
  const resp = await fetch("/api/chat", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6", max_tokens: 250,
      tools: [{ type: "web_search_20250305", name: "web_search" }],
      system: `Assistente magazzino bici. Dato un EAN trova il prodotto.
Rispondi SOLO in JSON senza markdown:
{"found":true|false,"brand":"","description":"<max 80 char in italiano>","category":""}`,
      messages: [{ role: "user", content: `EAN: ${ean}. Trova marca, descrizione breve e categoria ricambio bici.` }],
    }),
  });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`); // BUG FIX B (web lookup)
  const d = await resp.json();
  const raw = d.content?.filter(b => b.type === "text").map(b => b.text).join("") || "{}";
  const match = raw.match(/\{[\s\S]*\}/);
  if (!match) return { found: false };
  return JSON.parse(match[0]);
}

// ─── Excel export (ExcelJS — no vulnerabilities) ─────────────────────────────
async function exportToExcel(rows, includePrice) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "StockScan AI";
  workbook.created = new Date();
  const ws = workbook.addWorksheet("Inventario");

  const headers = ["Data","Ora","SKU","EAN","Descrizione","Scaffale","Q.tà AI","Confidenza","Note","Operatore"];
  if (includePrice) headers.push("Prezzo Acq. (€)");

  // Header row with bold + background
  const hdrRow = ws.addRow(headers);
  hdrRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  hdrRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1A1A1A" } };
  hdrRow.height = 16;

  // Column widths
  const widths = [11,8,15,16,32,10,8,12,30,13,...(includePrice?[15]:[])];
  ws.columns = headers.map((h, i) => ({ header: h, width: widths[i] || 12 }));

  // Data rows
  for (const r of rows) {
    const row = [r.date,r.time,r.sku||"—",r.ean||"—",r.desc||"—",r.shelf||"—",r.count,r.confidence,r.note||"",r.operator];
    if (includePrice) row.push(r.price != null ? r.price : "");
    const dataRow = ws.addRow(row);
    // Numeric price cell format
    if (includePrice && r.price != null) {
      dataRow.getCell(headers.length).numFmt = "#,##0.00";
    }
  }

  // Download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `inventario_${new Date().toISOString().slice(0,10)}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ─── Barcode Scanner Modal ────────────────────────────────────────────────────
function BarcodeModal({ onDetected, onClose }) {
  const videoRef = useRef();
  const streamRef = useRef();
  const [status, setStatus] = useState("init");
  const [errMsg, setErrMsg] = useState("");

  useEffect(() => {
    let interval;
    (async () => {
      try {
        const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } });
        streamRef.current = stream;
        if (!videoRef.current) return; // BUG FIX E: unmount prima di getUserMedia
        videoRef.current.srcObject = stream;
        await videoRef.current.play();
        if (!("BarcodeDetector" in window)) {
          setStatus("error");
          setErrMsg("BarcodeDetector non supportato. Usa Chrome su Android o Edge.");
          return;
        }
        const detector = new window.BarcodeDetector({ formats: ["ean_13","ean_8","code_128","code_39","upc_a","upc_e","qr_code"] });
        setStatus("scanning");
        interval = setInterval(async () => {
          try {
            const found = await detector.detect(videoRef.current);
            if (found.length > 0) {
              clearInterval(interval);
              setStatus("found");
              setTimeout(() => { stopStream(); onDetected(found[0].rawValue); }, 400);
            }
          } catch(_) {}
        }, 350);
      } catch(e) { setStatus("error"); setErrMsg("Camera non accessibile: " + e.message); }
    })();
    return () => { clearInterval(interval); stopStream(); };
    function stopStream() { streamRef.current?.getTracks().forEach(t => t.stop()); }
  }, []);

  return (
    <div style={{ position:"fixed", inset:0, background:"#000c", zIndex:50, display:"flex", alignItems:"center", justifyContent:"center" }}>
      <div style={{ width:"90%", maxWidth:400, background:C.surface, border:`1px solid ${C.border}`, borderRadius:2, overflow:"hidden" }}>
        <div style={{ padding:"8px 13px", background:C.surface2, borderBottom:`1px solid ${C.border}`, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <span style={{ fontSize:9, letterSpacing:2, color:C.orange, textTransform:"uppercase" }}>Scanner barcode</span>
          <button onClick={onClose} style={{ background:"none", border:"none", color:C.muted, cursor:"pointer", fontSize:18, lineHeight:1 }}>✕</button>
        </div>
        <div style={{ position:"relative", background:"#000", aspectRatio:"4/3" }}>
          <video ref={videoRef} style={{ width:"100%", height:"100%", objectFit:"cover", display:"block" }} muted playsInline />
          {status === "scanning" && (
            <div style={{ position:"absolute", inset:0, display:"flex", alignItems:"center", justifyContent:"center", pointerEvents:"none" }}>
              <div style={{ width:"70%", height:"28%", border:`2px solid ${C.orange}`, borderRadius:4, boxShadow:"0 0 0 9999px rgba(0,0,0,0.45)" }} />
            </div>
          )}
          {status === "found" && (
            <div style={{ position:"absolute", inset:0, background:"#00ff0033", display:"flex", alignItems:"center", justifyContent:"center" }}>
              <span style={{ color:C.green, fontSize:48 }}>✓</span>
            </div>
          )}
        </div>
        <div style={{ padding:"10px 13px", fontSize:10, color:status==="error"?C.red:C.muted, textAlign:"center" }}>
          {status==="error" ? errMsg : status==="found" ? "Codice rilevato!" : "Inquadra il codice a barre nell'area"}
        </div>
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function MagazzinoCounter() {
  const fileRef       = useRef();
  const anagraFileRef = useRef();

  const [imgSrc, setImgSrc]       = useState(null);
  const [imgBase64, setImgBase64] = useState(null);
  const [imgType, setImgType]     = useState("image/jpeg");

  const [sku, setSku] = useState("");
  const [ean, setEan] = useState("");

  const [desc, setDesc]         = useState("");
  const [price, setPrice]       = useState(null);
  const [shelf, setShelf]       = useState("");
  const [operator, setOperator] = useState("");
  const [descSource, setDescSource] = useState(null); // null | "anagrafica" | "web"

  const [eanWebLoading, setEanWebLoading] = useState(false);
  const [eanWebStatus, setEanWebStatus]   = useState(null); // null | "found" | "notfound"

  const [anagrafica, setAnagrafica] = useState(null);
  const [anaLoading, setAnaLoading] = useState(false);
  const [anaError, setAnaError]     = useState(null);

  const [scanTarget, setScanTarget]   = useState("ean");
  const [showScanner, setShowScanner] = useState(false);

  const [loading, setLoading] = useState(false);
  const [result, setResult]   = useState(null);
  const [error, setError]     = useState(null);

  const [log, setLog]                   = useState([]);
  const [includePrice, setIncludePrice] = useState(false);
  const [view, setView]                 = useState("scan");

  // ── Lookup in anagrafica by key ────────────────────────────────────────────
  const lookupAnag = useCallback((key) => {
    if (!anagrafica || !key) return null;
    return anagrafica.map[key.toUpperCase()] || anagrafica.map[key.replace(/\D/g,"")] || null;
  }, [anagrafica]);

  const applyEntry = useCallback((entry, { crossFillSku = false, crossFillEan = false } = {}) => {
    if (!entry) return;
    setDescSource("anagrafica"); // BUG FIX F: badge appare anche se desc è vuota
    if (entry.desc) setDesc(entry.desc);
    if (entry.price != null) setPrice(entry.price);
    if (crossFillSku && !sku && entry.sku) setSku(entry.sku);
    if (crossFillEan && !ean && entry.ean) setEan(entry.ean);
  }, [sku, ean]);

  // ── SKU field ──────────────────────────────────────────────────────────────
  const handleSkuChange = useCallback((val) => {
    const v = val.toUpperCase();
    setSku(v);
    setDescSource(null); setEanWebStatus(null);
    const entry = lookupAnag(v);
    if (entry) { applyEntry(entry, { crossFillEan: true }); return; }
    // BUG FIX C: se sku non trovato, mantieni desc/price se l'ean è ancora valido
    const fallback = lookupAnag(ean);
    if (fallback) { applyEntry(fallback); } else { setDesc(""); setPrice(null); }
  }, [lookupAnag, applyEntry, ean]);

  // ── EAN field ──────────────────────────────────────────────────────────────
  const handleEanChange = useCallback((val) => {
    const v = val.replace(/\D/g,"");
    setEan(v);
    setDescSource(null); setEanWebStatus(null);
    const entry = lookupAnag(v);
    if (entry) { applyEntry(entry, { crossFillSku: true }); return; }
    // BUG FIX C: se ean non trovato, mantieni desc/price se lo sku è ancora valido
    const fallback = lookupAnag(sku);
    if (fallback) { applyEntry(fallback); } else { setDesc(""); setPrice(null); }
  }, [lookupAnag, applyEntry, sku]);

  // ── Barcode scanner ────────────────────────────────────────────────────────
  const handleBarcodeDetected = useCallback((val) => {
    setShowScanner(false);
    if (scanTarget === "ean") handleEanChange(val);
    else handleSkuChange(val);
  }, [scanTarget, handleEanChange, handleSkuChange]);

  // ── EAN web lookup ─────────────────────────────────────────────────────────
  const handleWebLookup = useCallback(async () => {
    if (!ean || !isEAN(ean)) return;
    setEanWebLoading(true); setEanWebStatus(null);
    try {
      const res = await lookupEANweb(ean);
      if (res.found) {
        setDesc([res.brand, res.description].filter(Boolean).join(" — "));
        setDescSource("web"); setEanWebStatus("found");
      } else { setEanWebStatus("notfound"); }
    } catch { setEanWebStatus("notfound"); }
    finally { setEanWebLoading(false); }
  }, [ean]);

  // ── Camera ─────────────────────────────────────────────────────────────────
  const handleCapture = useCallback((e) => {
    const file = e.target.files[0]; if (!file) return;
    setResult(null); setError(null);
    setImgType(file.type || "image/jpeg");
    const reader = new FileReader();
    reader.onload = (ev) => { setImgSrc(ev.target.result); setImgBase64(ev.target.result.split(",")[1]); };
    reader.readAsDataURL(file);
  }, []);

  // ── Anagrafica ─────────────────────────────────────────────────────────────
  const handleAnagrafica = useCallback(async (e) => {
    const file = e.target.files[0]; if (!file) return;
    setAnaLoading(true); setAnaError(null);
    try { setAnagrafica(await parseAnagrafica(file)); }
    catch (err) { setAnaError(err.message); }
    finally {
      setAnaLoading(false);
      if (anagraFileRef.current) anagraFileRef.current.value = ""; // BUG FIX G
    }
  }, []);

  // ── AI count ───────────────────────────────────────────────────────────────
  const handleAnalyze = useCallback(async () => {
    if (!imgBase64) { setError("Scatta o carica prima una foto."); return; }
    if (!sku && !ean) { setError("Inserisci almeno SKU o EAN."); return; }
    setLoading(true); setError(null); setResult(null);
    try {
      const res = await countPiecesWithAI(imgBase64, imgType, desc || sku || ean);
      setResult(res);
      if (res.count >= 0) {
        const now = new Date();
        setLog(prev => [{
          id: Date.now(),
          date: now.toLocaleDateString("it-IT"),
          time: now.toLocaleTimeString("it-IT", { hour:"2-digit", minute:"2-digit" }),
          sku: sku||"—", ean: ean||"—", desc: desc||"—",
          shelf: shelf||"—", count: res.count,
          confidence: res.confidence, note: res.note||"",
          operator: operator||"—", price,
        }, ...prev]);
      }
    } catch(err) { setError("Errore API: " + err.message); }
    finally { setLoading(false); }
  }, [imgBase64, imgType, sku, ean, desc, shelf, operator, price]);

  const resetScan = () => {
    setImgSrc(null); setImgBase64(null); setResult(null); setError(null);
    setSku(""); setEan(""); setDesc(""); setPrice(null); setShelf("");
    setDescSource(null); setEanWebStatus(null);
    // BUG FIX 2: reset input file altrimenti stessa foto non ri-scatta onChange
    if (fileRef.current) fileRef.current.value = "";
  };

  const confColor = (c) => c === "alta" ? C.green : c === "media" ? C.yellow : C.red;

  // BUG FIX 4: useMemo evita il ricalcolo ad ogni render (potenzialmente
  // migliaia di articoli ad ogni keystroke).
  const anagStats = useMemo(() => {
    if (!anagrafica) return null;
    const seen = new Set(); const arr = [];
    for (const e of Object.values(anagrafica.map)) {
      const k = (e.sku||"") + "|" + (e.ean||"");
      if (!seen.has(k)) { seen.add(k); arr.push(e); }
    }
    return { total: arr.length, withSku: arr.filter(e=>e.sku).length, withEan: arr.filter(e=>e.ean).length, withPrice: arr.filter(e=>e.price!=null).length };
  }, [anagrafica]);

  return (
    <div style={S.app}>

      {showScanner && <BarcodeModal onDetected={handleBarcodeDetected} onClose={() => setShowScanner(false)} />}

      {/* HEADER */}
      <div style={S.header}>
        <div style={S.logo}>▦</div>
        <div style={{ flex:1 }}>
          <div style={{ fontSize:12, fontWeight:"bold", letterSpacing:2, color:C.orange, textTransform:"uppercase" }}>StockScan AI</div>
          <div style={{ fontSize:9, color:C.muted, letterSpacing:1 }}>Inventario ottico · ricambi bici</div>
        </div>
        <div style={{ textAlign:"right", fontSize:9, color:C.muted, lineHeight:1.7 }}>
          <span style={{ color:C.green }}>●</span> {log.length} scan
          {anagrafica && <><br /><span style={{ color:C.blue }}>▦</span> {anagStats.total.toLocaleString()} art.</>}
        </div>
      </div>

      {/* TABS */}
      <div style={{ display:"flex", background:"#111", borderBottom:`1px solid ${C.border}` }}>
        {[["scan","📷 Scan"],["log",`📋 Log (${log.length})`],["anagrafica","🗂 Anagrafica"]].map(([v,label]) => (
          <button key={v} onClick={() => setView(v)} style={{
            flex:1, padding:"9px 0", background: view===v ? C.surface : "transparent",
            border:"none", borderBottom: view===v ? `2px solid ${C.orange}` : "2px solid transparent",
            color: view===v ? C.orange : C.muted, fontFamily:"inherit",
            fontSize:9, letterSpacing:1.5, textTransform:"uppercase", cursor:"pointer",
          }}>{label}</button>
        ))}
      </div>

      <div style={S.body}>

        {/* ══ SCAN ══════════════════════════════════════════════════════════ */}
        {view === "scan" && (<>

          {/* 01 Photo */}
          <div style={S.card}>
            <div style={S.cardHead}>01 / Foto cassettone</div>
            <input ref={fileRef} type="file" accept="image/*" capture="environment" onChange={handleCapture} style={{ display:"none" }} />
            {imgSrc
              ? <img src={imgSrc} alt="" style={S.imgPreview} onClick={() => fileRef.current.click()} />
              : <div style={S.imgPlaceholder} onClick={() => fileRef.current.click()}>
                  <div style={{ fontSize:36, marginBottom:8 }}>📷</div>
                  <div style={{ fontSize:10, color:C.muted, letterSpacing:1 }}>TOCCA PER SCATTARE</div>
                  <div style={{ fontSize:9, color:"#333", marginTop:3 }}>o carica da galleria</div>
                </div>
            }
            {imgSrc && <div style={{ padding:"7px 13px" }}>
              <button style={S.btn("ghost","sm")} onClick={() => fileRef.current.click()}>↺ Nuova foto</button>
            </div>}
          </div>

          {/* 02 Article */}
          <div style={S.card}>
            <div style={S.cardHead}>02 / Identificazione articolo</div>
            <div style={S.cardBody}>

              {/* Source badge */}
              {descSource && (
                <div style={{ marginBottom:10, padding:"5px 10px", background: descSource==="anagrafica" ? "#0d2016" : "#0a1020", border:`1px solid ${descSource==="anagrafica" ? C.green : C.blue}33`, borderRadius:2, fontSize:9, color: descSource==="anagrafica" ? C.green : C.blue }}>
                  {descSource==="anagrafica" ? "✓ Articolo trovato in anagrafica" : "✓ Descrizione trovata online"}
                </div>
              )}

              {/* SKU */}
              <div style={{ marginBottom:10 }}>
                <label style={S.label}>Codice SKU interno <span style={{ color:C.muted }}>(opz.)</span></label>
                <div style={{ display:"flex", gap:7 }}>
                  <input
                    style={{ ...S.input(descSource==="anagrafica" && sku ? C.green+"66" : C.border), flex:1 }}
                    value={sku} onChange={e => handleSkuChange(e.target.value)}
                    placeholder="es. CHN-105-11V"
                  />
                  <button
                    title="Scansiona SKU con camera"
                    style={{ ...S.btn("ghost","sm"), width:"auto", padding:"8px 11px", flexShrink:0, color:C.orange, fontSize:14 }}
                    onClick={() => { setScanTarget("sku"); setShowScanner(true); }}
                  >⊡</button>
                </div>
              </div>

              {/* EAN */}
              <div style={{ marginBottom:10 }}>
                <label style={S.label}>Codice EAN <span style={{ color:C.muted }}>(opz.)</span></label>
                <div style={{ display:"flex", gap:7 }}>
                  <input
                    style={{ ...S.input(eanWebStatus==="found" ? C.blue+"66" : descSource==="anagrafica" && ean ? C.green+"66" : C.border), flex:1 }}
                    value={ean} onChange={e => handleEanChange(e.target.value)}
                    placeholder="es. 4550170439781"
                    inputMode="numeric" maxLength={14}
                  />
                  <button
                    title="Scansiona EAN con camera"
                    style={{ ...S.btn("ghost","sm"), width:"auto", padding:"8px 11px", flexShrink:0, color:C.orange, fontSize:14 }}
                    onClick={() => { setScanTarget("ean"); setShowScanner(true); }}
                  >⊡</button>
                  {ean.length >= 8 && isEAN(ean) && descSource !== "anagrafica" && (
                    <button
                      style={{ ...S.btn("ghost","sm"), width:"auto", padding:"6px 10px", flexShrink:0, color: eanWebLoading ? C.muted : C.orange, opacity: eanWebLoading ? 0.5 : 1, fontSize:9 }}
                      onClick={handleWebLookup} disabled={eanWebLoading}
                    >{eanWebLoading ? "⏳" : "🔍"}</button>
                  )}
                </div>
                {eanWebStatus==="found"    && <div style={{ marginTop:4, fontSize:9, color:C.blue }}>✓ Trovato online</div>}
                {eanWebStatus==="notfound" && <div style={{ marginTop:4, fontSize:9, color:C.red }}>✗ Non trovato — inserisci manualmente</div>}
              </div>

              {/* Description */}
              <div style={{ marginBottom:10 }}>
                <label style={S.label}>Descrizione <span style={{ color:C.muted }}>(opz. — auto da anagrafica / EAN)</span></label>
                <input
                  style={S.input(descSource ? (descSource==="anagrafica" ? C.green+"55" : C.blue+"55") : C.border)}
                  value={desc} onChange={e => setDesc(e.target.value)}
                  placeholder="auto-compilata o inserisci manualmente"
                />
              </div>

              {/* Price pill */}
              {price != null && (
                <div style={{ marginBottom:10, padding:"7px 11px", background:"#0d1a10", border:`1px solid ${C.green}33`, borderRadius:2, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                  <span style={{ fontSize:9, color:C.muted, letterSpacing:1 }}>PREZZO ACQUISTO (da anagrafica)</span>
                  <span style={{ color:C.green, fontWeight:"bold", fontSize:13 }}>€ {fmtPrice(price)}</span>
                </div>
              )}

              {/* Shelf + Operator */}
              <div style={S.row}>
                <div style={S.col}>
                  <label style={S.label}>Scaffale <span style={{ color:C.muted }}>(opz.)</span></label>
                  <input style={S.input()} value={shelf} onChange={e => setShelf(e.target.value.toUpperCase())} placeholder="es. A3-04" />
                </div>
                <div style={S.col}>
                  <label style={S.label}>Operatore</label>
                  <input style={S.input()} value={operator} onChange={e => setOperator(e.target.value)} placeholder="es. Marco" />
                </div>
              </div>
            </div>
          </div>

          {/* 03 AI */}
          <div style={S.card}>
            <div style={S.cardHead}>03 / Conteggio AI</div>
            <div style={S.cardBody}>
              <button style={{ ...S.btn("primary"), opacity: loading||!imgBase64 ? 0.4 : 1 }}
                onClick={handleAnalyze} disabled={loading||!imgBase64}>
                {loading ? "⏳ Analisi in corso..." : "⚡ CONTA CON AI"}
              </button>
              {error && <div style={{ marginTop:9, padding:"8px 11px", background:"#1a0808", border:`1px solid ${C.red}`, borderRadius:2, fontSize:10, color:C.red }}>⚠ {error}</div>}
              {result && (
                <div style={{ marginTop:10, padding:13, background: result.count>=0 ? "#0d2016" : "#1a0808", border:`1px solid ${result.count>=0 ? C.green : C.red}`, borderRadius:2 }}>
                  <div style={{ fontSize:9, color:C.muted, letterSpacing:1.5, marginBottom:4 }}>RISULTATO</div>
                  <div style={{ fontSize:36, fontWeight:"bold", color: result.count>=0 ? C.green : C.red, lineHeight:1 }}>
                    {result.count>=0 ? result.count : "—"}
                    <span style={{ fontSize:11, color:C.muted, marginLeft:7 }}>pz</span>
                  </div>
                  <div style={{ marginTop:8, display:"flex", gap:7, flexWrap:"wrap", alignItems:"center" }}>
                    <span style={S.tag(confColor(result.confidence))}>{result.confidence?.toUpperCase()}</span>
                    {result.note && <span style={{ fontSize:9, color:C.muted }}>{result.note}</span>}
                  </div>
                  <div style={{ marginTop:11 }}>
                    <button style={S.btn("ghost","sm")} onClick={resetScan}>＋ Nuova scansione</button>
                  </div>
                </div>
              )}
            </div>
          </div>
        </>)}

        {/* ══ LOG ═══════════════════════════════════════════════════════════ */}
        {view === "log" && (<>
          {log.length > 0 && (
            <div style={S.card}>
              <div style={S.cardHead}>Opzioni export</div>
              <div style={{ padding:"11px 13px" }}>
                <label style={{ display:"flex", alignItems:"center", gap:10, cursor:"pointer", fontSize:11 }}>
                  <input type="checkbox" checked={includePrice} onChange={e => setIncludePrice(e.target.checked)}
                    style={{ accentColor:C.orange, width:15, height:15 }} />
                  <span>Includi <strong>Prezzo Acquisto</strong> nell'Excel</span>
                </label>
              </div>
            </div>
          )}
          {log.length > 0 && (
            <div style={{ marginBottom:12 }}>
              <button style={S.btn("primary")} onClick={() => exportToExcel(log, includePrice).catch(err => alert("Errore export: " + err.message))}>⬇ ESPORTA EXCEL (.xlsx)</button>
            </div>
          )}
          <div style={S.card}>
            <div style={S.cardHead}>Storico sessione</div>
            {log.length === 0
              ? <div style={{ padding:"28px", textAlign:"center", color:C.muted, fontSize:11 }}>Nessuna scansione.<br /><span style={{ color:"#333" }}>Vai su Scan per iniziare.</span></div>
              : <div style={{ overflowX:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
                    <thead><tr>
                      <th style={S.th}>Ora</th>
                      <th style={S.th}>SKU</th>
                      <th style={S.th}>EAN</th>
                      <th style={S.th}>Q.</th>
                      <th style={S.th}>Conf.</th>
                      {includePrice && <th style={S.th}>€</th>}
                    </tr></thead>
                    <tbody>
                      {log.map(r => (
                        <tr key={r.id}>
                          <td style={{ ...S.td, color:C.muted, fontSize:10 }}>{r.time}</td>
                          <td style={{ ...S.td, color:C.orange, fontWeight:"bold", fontSize:11 }}>{r.sku}</td>
                          <td style={{ ...S.td, color:C.muted, fontSize:10 }}>{r.ean}</td>
                          <td style={{ ...S.td, color:C.green, fontWeight:"bold", fontSize:14 }}>{r.count}</td>
                          <td style={S.td}><span style={S.tag(confColor(r.confidence))}>{r.confidence}</span></td>
                          {includePrice && <td style={{ ...S.td, color:C.green, fontSize:10 }}>{r.price!=null ? `€${fmtPrice(r.price)}` : "—"}</td>}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
            }
            {log.length > 0 && (
              <div style={{ padding:"9px 13px", fontSize:10, color:C.muted, borderTop:`1px solid ${C.border}`, display:"flex", gap:8, alignItems:"center" }}>
                <span style={{ width:6, height:6, borderRadius:"50%", background:C.green, display:"inline-block" }} />
                {log.length} art. · {log.reduce((s,r) => s+(r.count||0), 0)} pz totali
              </div>
            )}
          </div>
          {log.length > 0 && (
            <button style={S.btn("danger","sm")} onClick={() => { if(confirm("Cancellare tutto il log?")) setLog([]); }}>
              ✕ Cancella log sessione
            </button>
          )}
        </>)}

        {/* ══ ANAGRAFICA ════════════════════════════════════════════════════ */}
        {view === "anagrafica" && (<>
          <div style={S.card}>
            <div style={S.cardHead}>Carica listino / anagrafica</div>
            <div style={S.cardBody}>
              <div style={{ fontSize:10, color:C.muted, lineHeight:1.8, marginBottom:12 }}>
                Carica un file <strong style={{ color:C.text }}>CSV o XLSX</strong> esportato da Ad Hoc Revolution Web o qualsiasi gestionale.<br />
                Colonne rilevate automaticamente:
                <div style={{ marginTop:6, display:"flex", gap:6, flexWrap:"wrap" }}>
                  {["SKU / Codice art.","EAN / Barcode","Descrizione","Prezzo acquisto"].map(c => (
                    <span key={c} style={{ ...S.tag(C.orange), fontSize:9 }}>{c}</span>
                  ))}
                </div>
              </div>
              <input ref={anagraFileRef} type="file" accept=".csv,.xlsx,.xls" onChange={handleAnagrafica} style={{ display:"none" }} />
              <button style={{ ...S.btn("primary"), opacity: anaLoading ? 0.5 : 1 }}
                onClick={() => anagraFileRef.current.click()} disabled={anaLoading}>
                {anaLoading ? "⏳ Caricamento..." : "📂 CARICA FILE CSV / XLSX"}
              </button>
              {anaError && <div style={{ marginTop:9, padding:"8px 11px", background:"#1a0808", border:`1px solid ${C.red}`, borderRadius:2, fontSize:10, color:C.red }}>⚠ {anaError}</div>}
            </div>
          </div>

          {anagrafica && anagStats && (
            <div style={S.card}>
              <div style={S.cardHead}>✓ Anagrafica caricata</div>
              <div style={S.cardBody}>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:9, marginBottom:12 }}>
                  {[["Articoli totali", anagStats.total.toLocaleString(), C.text],
                    ["Con SKU", anagStats.withSku.toLocaleString(), C.orange],
                    ["Con EAN", anagStats.withEan.toLocaleString(), C.blue],
                    ["Con prezzo", anagStats.withPrice.toLocaleString(), C.green],
                  ].map(([label, val, color]) => (
                    <div key={label} style={{ padding:"9px 11px", background:"#111", border:`1px solid ${C.border}`, borderRadius:2 }}>
                      <div style={{ fontSize:9, color:C.muted, letterSpacing:1, marginBottom:3 }}>{label.toUpperCase()}</div>
                      <div style={{ fontSize:18, fontWeight:"bold", color }}>{val}</div>
                    </div>
                  ))}
                </div>
                <button style={S.btn("ghost","sm")} onClick={() => { setAnagrafica(null); anagraFileRef.current.value=""; }}>
                  ✕ Rimuovi anagrafica
                </button>
              </div>
            </div>
          )}

          {!anagrafica && (
            <div style={{ padding:"16px 13px", fontSize:10, color:C.muted, lineHeight:1.8 }}>
              <strong style={{ color:C.text, display:"block", marginBottom:6 }}>Come funziona</strong>
              Una volta caricata l'anagrafica, quando inserisci o scansioni uno dei codici nella schermata Scan, la descrizione e il prezzo vengono compilati automaticamente.<br /><br />
              Se l'articolo non è in anagrafica e hai un EAN, puoi cercare la descrizione online con il tasto 🔍.
            </div>
          )}
        </>)}

      </div>
    </div>
  );
}
