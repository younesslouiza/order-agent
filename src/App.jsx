import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
const XLSX = window.XLSX;

const COLS = {
  CT: "CT-Number",
  EAN: "EAN",
  TITLE: "Title",
  MANUFACTURER: "Manufacturer",
  VARIANT: "Variant",
  MIN_STOCK: "Minimum Stock",
  STOCK_S: "Current Stock (Schulstrasse)",
  STOCK_L: "Current Stock (Laufenburg)",
  PURCHASE_PRICE: "Purchase Price",
  SALES_MONTH: "Sales Per Month",
};

const PERIODS = [
  { label: "1W",  value: 1/4,  display: "1 Week" },
  { label: "2W",  value: 2/4,  display: "2 Weeks" },
  { label: "1M",  value: 1,    display: "1 Month" },
  { label: "2M",  value: 2,    display: "2 Months" },
  { label: "3M",  value: 3,    display: "3 Months" },
  { label: "6M",  value: 6,    display: "6 Months" },
  { label: "12M", value: 12,   display: "12 Months" },
];

// Detect file type from columns
function detectWarehouse(columns) {
  const hasNonPOS = columns.some((c) => c.includes("Non-POS Sales"));
  const hasPOS = columns.some((c) => c.includes("POS Sales") && !c.includes("Non-POS"));
  const hasDecision = columns.some((c) => c.includes("Decision"));
  if (hasNonPOS) return "laufenburg";
  if (hasPOS || hasDecision) return "schulstrasse";
  return "laufenburg";
}

function getSalesYearCol(columns) {
  return columns.find((c) => c.includes("POS Sales") || c.includes("Non-POS Sales")) || "";
}

function getSalesMonthCol(columns) {
  return columns.find((c) =>
    c.toLowerCase().includes("sale per month") ||
    c.toLowerCase().includes("sales per month")
  ) || "Sales Per Month";
}

function getDemandCol(columns) {
  return columns.find((c) =>
    c.toLowerCase().includes("demand") ||
    c.toLowerCase().includes("deamnd")
  ) || "Demand for X Month";
}

// Schulstrasse: Desired = Demand - Stock_S, can order from Warehouse OR Supplier
function computeSchulstrasse(row, periodValue, salesMonthCol) {
  const salesMonth = Number(row[salesMonthCol] || row[COLS.SALES_MONTH]) || 0;
  const stockS = Number(row[COLS.STOCK_S]) || 0;
  const stockL = Number(row[COLS.STOCK_L]) || 0;
  const purchasePrice = Number(row[COLS.PURCHASE_PRICE]) || 0;

  const demand = Math.ceil(salesMonth * periodValue);
  const desired = Math.max(0, demand - stockS);

  let orderW = 0, orderS = 0, decision = "No Order Needed";
  if (desired > 0) {
    orderW = Math.min(Math.floor(stockL / 6), desired);
    orderS = Math.max(0, desired - orderW);
    if (orderW > 0 && orderS > 0) decision = `Split: ${orderW} Warehouse + ${orderS} Supplier`;
    else if (orderW > 0) decision = "All from Warehouse";
    else decision = "All from Supplier";
  }

  return { ...row, _demand: demand, _desired: desired, _orderW: orderW, _orderS: orderS, _decision: decision, _totalPrice: (orderS * purchasePrice).toFixed(2) };
}

// Laufenburg: Desired = Demand - Stock_L, ONLY from Supplier
function computeLaufenburg(row, periodValue, salesMonthCol) {
  const salesMonth = Number(row[salesMonthCol] || row[COLS.SALES_MONTH]) || 0;
  const stockL = Number(row[COLS.STOCK_L]) || 0;
  const purchasePrice = Number(row[COLS.PURCHASE_PRICE]) || 0;

  const demand = Math.ceil(salesMonth * periodValue);
  const desired = Math.max(0, demand - stockL);

  let orderS = 0, decision = "No Order Needed";
  if (desired > 0) {
    orderS = desired;
    decision = "All from Supplier";
  }

  return { ...row, _demand: demand, _desired: desired, _orderW: 0, _orderS: orderS, _decision: decision, _totalPrice: (orderS * purchasePrice).toFixed(2) };
}

function isOrderable(row, salesCol) {
  return (Number(row[COLS.MIN_STOCK]) || 0) !== 0 && (Number(row[salesCol]) || 0) !== 0;
}

function isNewProduct(row, salesCol) {
  const creationDate = row["Creation Date"];
  let year = null;
  if (creationDate) {
    const d = new Date(creationDate);
    if (!isNaN(d)) year = d.getFullYear();
    else {
      const m = String(creationDate).match(/(\d{4})/);
      if (m) year = parseInt(m[1]);
    }
  }
  return (
    (Number(row[COLS.MIN_STOCK]) || 0) !== 0 &&
    (Number(row[COLS.STOCK_S]) || 0) === 0 &&
    (Number(row[COLS.STOCK_L]) || 0) === 0 &&
    (Number(row[salesCol]) || 0) === 0 &&
    (year === 2025 || year === 2026)
  );
}

function buildNewProductSheet(rows) {
  const exportData = rows.map((r) => ({
    "CT-Number": r[COLS.CT] || "",
    "EAN": String(r[COLS.EAN] || ""),
    "Manufacturer Number": String(r["Manufacturer Number"] || ""),
    "Manufacturer": r[COLS.MANUFACTURER] || "",
    "Title": r[COLS.TITLE] || "",
    "Variant": r[COLS.VARIANT] || "",
    "Creation Date": r["Creation Date"] || "",
    "Desired Quantity": "New",
  }));
  const ws = XLSX.utils.json_to_sheet(exportData);
  ws["!cols"] = [{ wch: 14 }, { wch: 16 }, { wch: 20 }, { wch: 28 }, { wch: 50 }, { wch: 30 }, { wch: 14 }, { wch: 12 }];
  return ws;
}

function buildSheet(rows, qtyKey) {
  const exportData = rows.map((r) => ({
    "CT-Number": r[COLS.CT] || "",
    "EAN": String(r[COLS.EAN] || ""),
    "Manufacturer Number": String(r["Manufacturer Number"] || ""),
    "Manufacturer": r[COLS.MANUFACTURER] || "",
    "Title": r[COLS.TITLE] || "",
    "Variant": r[COLS.VARIANT] || "",
    "Desired Quantity": qtyKey === "_orderW" ? r._orderW : qtyKey === "_orderS" ? r._orderS : r._desired,
  }));
  const ws = XLSX.utils.json_to_sheet(exportData);
  ws["!cols"] = [{ wch: 14 }, { wch: 16 }, { wch: 20 }, { wch: 28 }, { wch: 50 }, { wch: 30 }, { wch: 16 }];
  return ws;
}

function downloadExcel(rows, periodDisplay, warehouse) {
  const date = new Date().toISOString().slice(0, 10);
  const per = periodDisplay.replace(" ", "_");

  if (warehouse === "schulstrasse") {
    // Warehouse file
    const warehouseRows = rows.filter((r) => r._orderW > 0);
    if (warehouseRows.length > 0) {
      const wb1 = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb1, buildSheet(warehouseRows, "_orderW"), "Warehouse");
      XLSX.writeFile(wb1, `order_warehouse_${per}_${date}.xlsx`);
    }
    // Supplier file
    const supplierRows = rows.filter((r) => r._orderS > 0);
    if (supplierRows.length > 0) {
      const wb2 = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb2, buildSheet(supplierRows, "_orderS"), "Supplier");
      XLSX.writeFile(wb2, `order_supplier_${per}_${date}.xlsx`);
    }
  } else {
    // Laufenburg — single file, all from supplier
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, buildSheet(rows, "_desired"), "Supplier");
    XLSX.writeFile(wb, `order_laufenburg_supplier_${per}_${date}.xlsx`);
  }
}

export default function App() {
  const [rawData, setRawData] = useState(null);
  const [salesCol, setSalesCol] = useState("");
  const [salesMonthColName, setSalesMonthColName] = useState("Sales Per Month");
  const [warehouse, setWarehouse] = useState(""); // "schulstrasse" | "laufenburg"
  const [fileName, setFileName] = useState("");
  const [periodIdx, setPeriodIdx] = useState(2);
  const [filterDecision, setFilterDecision] = useState("all");
  const [filterBrand, setFilterBrand] = useState("all");
  const [search, setSearch] = useState("");
  const [loading, setLoading] = useState(false);

  const period = PERIODS[periodIdx];

  const handleFile = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);
    setFileName(file.name);
    setFilterDecision("all");
    setFilterBrand("all");
    setSearch("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws);
      if (!rows.length) { setLoading(false); return; }
      const cols = Object.keys(rows[0]);
      const sc = getSalesYearCol(cols);
      const wh = detectWarehouse(cols);
      const smc = getSalesMonthCol(cols);
      setSalesCol(sc);
      setWarehouse(wh);
      setSalesMonthColName(smc);
      setRawData(rows);
      setLoading(false);
    };
    reader.readAsBinaryString(file);
  }, []);

  const processed = useMemo(() => {
    if (!rawData || !salesCol) return [];
    return rawData
      .filter((r) => isOrderable(r, salesCol))
      .map((r) => warehouse === "schulstrasse"
        ? computeSchulstrasse(r, period.value, salesMonthColName)
        : computeLaufenburg(r, period.value, salesMonthColName)
      );
  }, [rawData, salesCol, warehouse, periodIdx, salesMonthColName]);

  const orderable = useMemo(() => processed.filter((r) => r._desired > 0), [processed]);

  const newProducts = useMemo(() => {
    if (!rawData || !salesCol || warehouse !== "schulstrasse") return [];
    return rawData.filter((r) => isNewProduct(r, salesCol));
  }, [rawData, salesCol, warehouse]);

  const brands = useMemo(() =>
    ["all", ...Array.from(new Set(orderable.map((r) => r[COLS.MANUFACTURER]).filter(Boolean))).sort()],
    [orderable]
  );

  const filtered = useMemo(() => orderable.filter((row) => {
    const d = row._decision || "";
    const matchD =
      filterDecision === "all" ||
      (filterDecision === "supplier" && d === "All from Supplier") ||
      (filterDecision === "warehouse" && d === "All from Warehouse") ||
      (filterDecision === "split" && d.startsWith("Split"));
    const matchB = filterBrand === "all" || row[COLS.MANUFACTURER] === filterBrand;
    const matchS = !search ||
      (row[COLS.TITLE] || "").toLowerCase().includes(search.toLowerCase()) ||
      (row[COLS.CT] || "").toLowerCase().includes(search.toLowerCase()) ||
      (row[COLS.MANUFACTURER] || "").toLowerCase().includes(search.toLowerCase());
    return matchD && matchB && matchS;
  }), [orderable, filterDecision, filterBrand, search]);

  const stats = useMemo(() => ({
    total: orderable.length,
    supplier: orderable.filter((r) => r._decision === "All from Supplier").length,
    warehouse: orderable.filter((r) => r._decision === "All from Warehouse").length,
    split: orderable.filter((r) => r._decision?.startsWith("Split")).length,
    totalCost: orderable.reduce((s, r) => s + (Number(r._totalPrice) || 0), 0).toFixed(2),
  }), [orderable]);

  const filteredSupplierTotal = useMemo(() =>
    filtered.reduce((s, r) => s + (r._orderS * (Number(r[COLS.PURCHASE_PRICE]) || 0)), 0).toFixed(2),
    [filtered]
  );

  const isSchulstrasse = warehouse === "schulstrasse";
  const warehouseColor = isSchulstrasse ? "#8b5cf6" : "#f59e0b";
  const fileBaseName = fileName.replace(/\.(xlsx|xls)$/i, "");
  const warehouseLabel = (isSchulstrasse ? "🟣 " : "🟡 ") + fileBaseName;

  const decisionColor = (d) => d === "All from Warehouse" ? "#10b981" : d?.startsWith("Split") ? "#f59e0b" : "#3b82f6";
  const decisionBg   = (d) => d === "All from Warehouse" ? "rgba(16,185,129,0.12)" : d?.startsWith("Split") ? "rgba(245,158,11,0.12)" : "rgba(59,130,246,0.12)";
  const inp = { background: "#141820", border: "1px solid #1e2530", borderRadius: 8, color: "#e2e8f0", outline: "none" };

  return (
    <div style={{ minHeight: "100vh", background: "#0f1117", color: "#e2e8f0", fontFamily: "'DM Mono','Courier New',monospace" }}>

      {/* Header */}
      <div style={{ borderBottom: "1px solid #1e2530", padding: "16px 28px", display: "flex", alignItems: "center", justifyContent: "space-between", background: "#0a0d13" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <div style={{ width: 34, height: 34, background: "linear-gradient(135deg,#3b82f6,#8b5cf6)", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>⚡</div>
          <div>
            <div style={{ fontSize: 16, fontWeight: 700, color: "#f1f5f9" }}>C-Total Order Agent</div>
            <div style={{ fontSize: 10, color: "#64748b", letterSpacing: "1px", textTransform: "uppercase" }}>Smart Purchase Decision Engine</div>
          </div>
        </div>
        {rawData && (
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            {/* Warehouse + filename badge */}
            <div style={{ padding: "5px 18px", borderRadius: 20, background: `${warehouseColor}22`, border: `1px solid ${warehouseColor}55`, color: warehouseColor, fontSize: 13, fontWeight: 700, letterSpacing: "0.3px", maxWidth: 320, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {warehouseLabel}
            </div>
            <label style={{ padding: "7px 14px", ...inp, cursor: "pointer", fontSize: 11, color: "#94a3b8" }}>
              Change File <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
            </label>
          </div>
        )}
      </div>

      <div style={{ padding: "24px 28px" }}>

        {/* Upload */}
        {!rawData && (
          <div style={{ maxWidth: 580, margin: "60px auto", textAlign: "center" }}>
            <div style={{ fontSize: 52, marginBottom: 16 }}>📊</div>
            <h2 style={{ fontSize: 22, fontWeight: 700, color: "#f1f5f9", marginBottom: 8 }}>Import Sales Excel</h2>
            <p style={{ color: "#64748b", marginBottom: 20, fontSize: 13, lineHeight: 1.7 }}>
              The agent automatically detects which warehouse file you upload<br />and applies the correct order logic.
            </p>

            {/* Two warehouse cards */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 28, textAlign: "left" }}>
              {[
                { color: "#8b5cf6", icon: "🟣", name: "Schulstrasse", rules: ["Min Stock ≠ 0", "POS Sales ≠ 0", "Desired > 0"], sources: ["🏭 Warehouse (Laufenburg)", "🚚 Supplier"] },
                { color: "#f59e0b", icon: "🟡", name: "Laufenburg",   rules: ["Min Stock ≠ 0", "Non-POS Sales ≠ 0", "Desired > 0"], sources: ["🚚 Supplier only"] },
              ].map((w) => (
                <div key={w.name} style={{ background: "#141820", border: `1px solid ${w.color}44`, borderRadius: 12, padding: "16px 18px" }}>
                  <div style={{ fontSize: 13, fontWeight: 700, color: w.color, marginBottom: 10 }}>{w.icon} {w.name}</div>
                  <div style={{ fontSize: 11, color: "#475569", marginBottom: 6, textTransform: "uppercase", letterSpacing: "0.5px" }}>Filters</div>
                  {w.rules.map((r) => <div key={r} style={{ fontSize: 11, color: "#64748b", marginBottom: 3 }}>✓ {r}</div>)}
                  <div style={{ fontSize: 11, color: "#475569", marginTop: 10, marginBottom: 6, textTransform: "uppercase", letterSpacing: "0.5px" }}>Orders from</div>
                  {w.sources.map((s) => <div key={s} style={{ fontSize: 11, color: "#94a3b8", marginBottom: 3 }}>{s}</div>)}
                </div>
              ))}
            </div>

            <label style={{ display: "inline-block", padding: "13px 30px", background: "linear-gradient(135deg,#3b82f6,#8b5cf6)", borderRadius: 10, cursor: "pointer", fontSize: 14, fontWeight: 600, color: "#fff" }}>
              {loading ? "⏳ Processing..." : "📁 Choose Excel File"}
              <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
            </label>
          </div>
        )}

        {/* Period Selector */}
        {rawData && (
          <div style={{ background: "#141820", border: "1px solid #1e2530", borderRadius: 12, padding: "16px 24px", marginBottom: 18, display: "flex", alignItems: "center", gap: 24, flexWrap: "wrap" }}>
            <div>
              <div style={{ fontSize: 10, color: "#475569", textTransform: "uppercase", letterSpacing: "0.5px", marginBottom: 8 }}>📅 Order Period</div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                {PERIODS.map((p, i) => (
                  <button key={p.label} onClick={() => setPeriodIdx(i)} style={{
                    padding: "8px 14px", borderRadius: 8, border: "1px solid",
                    borderColor: periodIdx === i ? warehouseColor : "#1e2530",
                    background: periodIdx === i ? `${warehouseColor}22` : "transparent",
                    color: periodIdx === i ? warehouseColor : "#475569",
                    cursor: "pointer", fontSize: 12, fontWeight: periodIdx === i ? 700 : 500,
                  }}>{p.label}</button>
                ))}
              </div>
            </div>
            <div style={{ height: 40, width: 1, background: "#1e2530" }} />
            <div style={{ fontSize: 12, color: "#64748b", lineHeight: 2 }}>
              <div>📐 <span style={{ color: "#94a3b8" }}>Demand</span> = Sales/Month × <span style={{ color: warehouseColor, fontWeight: 700 }}>{period.value}</span> <span style={{ color: "#374151" }}>({period.display})</span></div>
              <div>📦 <span style={{ color: "#94a3b8" }}>Desired</span> = Demand − Stock {isSchulstrasse ? "Schulstrasse" : "Laufenburg"}</div>
            </div>
            {/* Laufenburg notice */}
            {!isSchulstrasse && (
              <div style={{ marginLeft: "auto", padding: "8px 14px", background: "rgba(245,158,11,0.08)", border: "1px solid rgba(245,158,11,0.3)", borderRadius: 8, fontSize: 11, color: "#fbbf24" }}>
                🚚 Laufenburg orders from Supplier only
              </div>
            )}
          </div>
        )}

        {/* Stats */}
        {rawData && (
          <div style={{ display: "grid", gridTemplateColumns: isSchulstrasse ? "repeat(5,1fr)" : "repeat(3,1fr)", gap: 14, marginBottom: 18 }}>
            {[
              { label: "Products to Order", value: stats.total, icon: "📦", color: "#f1f5f9", show: true },
              { label: "From Supplier", value: stats.supplier, icon: "🚚", color: "#3b82f6", show: true },
              { label: "From Warehouse", value: stats.warehouse, icon: "🏭", color: "#10b981", show: isSchulstrasse },

              { label: "Total Cost CHF", value: Number(stats.totalCost).toLocaleString(), icon: "💰", color: "#a78bfa", show: true },
            ].filter((s) => s.show).map((s) => (
              <div key={s.label} style={{ background: "#141820", border: "1px solid #1e2530", borderRadius: 10, padding: "14px 16px" }}>
                <div style={{ fontSize: 18, marginBottom: 4 }}>{s.icon}</div>
                <div style={{ fontSize: 20, fontWeight: 700, color: s.color }}>{s.value}</div>
                <div style={{ fontSize: 10, color: "#475569", textTransform: "uppercase", letterSpacing: "0.5px", marginTop: 2 }}>{s.label}</div>
              </div>
            ))}
          </div>
        )}

        {/* Filters */}
        {rawData && (
          <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap", alignItems: "center" }}>
            <input placeholder="🔍  Search product, brand, CT-Number..." value={search} onChange={(e) => setSearch(e.target.value)}
              style={{ flex: 1, minWidth: 200, padding: "9px 13px", fontSize: 12, ...inp }} />

            <select value={filterBrand} onChange={(e) => setFilterBrand(e.target.value)}
              style={{ padding: "9px 13px", fontSize: 12, maxWidth: 200, ...inp }}>
              {brands.map((b) => <option key={b} value={b}>{b === "all" ? "All Brands" : b}</option>)}
            </select>

            {/* Decision filters — hide warehouse/split for Laufenburg */}
            {[
              { key: "all", label: `All (${orderable.length})`, show: true },
              { key: "supplier", label: `🚚 Supplier (${stats.supplier})`, show: true },
              { key: "warehouse", label: `🏭 Warehouse (${stats.warehouse})`, show: isSchulstrasse },

            ].filter((f) => f.show).map((f) => (
              <button key={f.key} onClick={() => setFilterDecision(f.key)} style={{
                padding: "8px 12px", borderRadius: 8, border: "1px solid",
                borderColor: filterDecision === f.key ? warehouseColor : "#1e2530",
                background: filterDecision === f.key ? `${warehouseColor}22` : "#141820",
                color: filterDecision === f.key ? warehouseColor : "#64748b",
                cursor: "pointer", fontSize: 11, fontWeight: 600,
              }}>{f.label}</button>
            ))}

            {isSchulstrasse ? (
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={() => downloadExcel(filtered.filter(r => r._orderW > 0), period.display, warehouse)}
                  style={{ padding: "9px 14px", borderRadius: 8, background: "linear-gradient(135deg,#10b981,#059669)", color: "#fff", border: "none", cursor: "pointer", fontSize: 11, fontWeight: 700 }}>
                  ⬇️ Warehouse ({filtered.filter(r => r._orderW > 0).length})
                </button>
                <button onClick={() => downloadExcel(filtered.filter(r => r._orderS > 0), period.display, warehouse)}
                  style={{ padding: "9px 14px", borderRadius: 8, background: "linear-gradient(135deg,#3b82f6,#1d4ed8)", color: "#fff", border: "none", cursor: "pointer", fontSize: 11, fontWeight: 700 }}>
                  ⬇️ Supplier ({filtered.filter(r => r._orderS > 0).length})
                </button>
              </div>
            ) : (
              <button onClick={() => downloadExcel(filtered, period.display, warehouse)}
                style={{ padding: "9px 14px", borderRadius: 8, background: "linear-gradient(135deg,#3b82f6,#1d4ed8)", color: "#fff", border: "none", cursor: "pointer", fontSize: 11, fontWeight: 700 }}>
                ⬇️ Supplier ({filtered.length})
              </button>
            )}
            {isSchulstrasse && newProducts.length > 0 && (
              <button onClick={() => {
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, buildNewProductSheet(newProducts), "New Products");
                XLSX.writeFile(wb, `new_products_${new Date().toISOString().slice(0,10)}.xlsx`);
              }}
                style={{ padding: "9px 14px", borderRadius: 8, background: "linear-gradient(135deg,#f59e0b,#d97706)", color: "#fff", border: "none", cursor: "pointer", fontSize: 11, fontWeight: 700 }}>
                ⭐ New Products ({newProducts.length})
              </button>
            )}
          </div>
        )}

        {/* Total CHF bar */}
        {rawData && filtered.length > 0 && (
          <div style={{ background: "rgba(167,139,250,0.06)", border: "1px solid rgba(167,139,250,0.2)", borderRadius: 8, padding: "10px 18px", marginBottom: 10, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <span style={{ fontSize: 12, color: "#94a3b8" }}>
              💰 Total Supplier Cost <span style={{ color: "#64748b", fontSize: 11 }}>({filtered.filter(r => r._orderS > 0).length} products × Purchase Price)</span>
            </span>
            <span style={{ fontSize: 18, fontWeight: 700, color: "#a78bfa", letterSpacing: "-0.5px" }}>
              CHF {Number(filteredSupplierTotal).toLocaleString("de-CH", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
            </span>
          </div>
        )}

        {/* Export info bar */}
        {rawData && filtered.length > 0 && (
          <div style={{ background: "rgba(16,185,129,0.06)", border: "1px solid rgba(16,185,129,0.2)", borderRadius: 8, padding: "9px 16px", marginBottom: 14, fontSize: 11, color: "#6ee7b7", display: "flex", alignItems: "center", gap: 8 }}>
            <span>📋</span>
            {isSchulstrasse ? (
              <span>
                🏭 <strong>Warehouse:</strong> {filtered.filter(r => r._orderW > 0).length} &nbsp;|&nbsp;
                🚚 <strong>Supplier:</strong> {filtered.filter(r => r._orderS > 0).length} &nbsp;|&nbsp;

                ⭐ <strong>New Products:</strong> {newProducts.length}
              </span>
            ) : (
              <span>🚚 <strong>Supplier:</strong> {filtered.length} products — {period.display} · {filterBrand !== "all" ? filterBrand : "All Brands"}</span>
            )}
          </div>
        )}

        {/* Table */}
        {rawData && (
          <div style={{ background: "#141820", border: "1px solid #1e2530", borderRadius: 12, overflow: "hidden" }}>
            <div style={{ overflowX: "auto", maxHeight: "48vh", overflowY: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: "#0a0d13", position: "sticky", top: 0, zIndex: 1 }}>
                    {[
                      "CT-Number","Manufacturer","Title","Variant",
                      "Min St.", isSchulstrasse ? "Stock S." : "Stock L.",
                      isSchulstrasse ? "Stock L." : null,
                      "Sales/Mo", `Demand(${period.label})`, "Desired",
                      isSchulstrasse ? "Warehouse" : null,
                      "Supplier","CHF","Decision"
                    ].filter(Boolean).map((h) => (
                      <th key={h} style={{ padding: "10px 12px", textAlign: "left", color: "#475569", fontWeight: 600, fontSize: 10, textTransform: "uppercase", letterSpacing: "0.5px", borderBottom: "1px solid #1e2530", whiteSpace: "nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.slice(0, 500).map((row, i) => {
                    const d = row._decision || "";
                    return (
                      <tr key={i} style={{ borderBottom: "1px solid #151c27", background: i % 2 === 0 ? "transparent" : "rgba(255,255,255,0.01)" }}>
                        <td style={{ padding: "8px 12px", color: "#64748b", fontFamily: "monospace", whiteSpace: "nowrap" }}>{row[COLS.CT]}</td>
                        <td style={{ padding: "8px 12px", color: "#94a3b8", whiteSpace: "nowrap" }}>{row[COLS.MANUFACTURER]}</td>
                        <td style={{ padding: "8px 12px", color: "#cbd5e1", maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={row[COLS.TITLE]}>{row[COLS.TITLE]}</td>
                        <td style={{ padding: "8px 12px", color: "#64748b", maxWidth: 110, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{row[COLS.VARIANT]}</td>
                        <td style={{ padding: "8px 12px", color: "#94a3b8", textAlign: "center" }}>{row[COLS.MIN_STOCK]}</td>
                        <td style={{ padding: "8px 12px", color: "#94a3b8", textAlign: "center" }}>{isSchulstrasse ? row[COLS.STOCK_S] : row[COLS.STOCK_L]}</td>
                        {isSchulstrasse && <td style={{ padding: "8px 12px", color: "#94a3b8", textAlign: "center" }}>{row[COLS.STOCK_L]}</td>}
                        <td style={{ padding: "8px 12px", color: "#94a3b8", textAlign: "center" }}>{row[COLS.SALES_MONTH]}</td>
                        <td style={{ padding: "8px 12px", color: "#a78bfa", fontWeight: 600, textAlign: "center" }}>{row._demand}</td>
                        <td style={{ padding: "8px 12px", fontWeight: 700, color: "#f1f5f9", textAlign: "center" }}>{row._desired}</td>
                        {isSchulstrasse && <td style={{ padding: "8px 12px", color: "#10b981", fontWeight: 700, textAlign: "center" }}>{row._orderW > 0 ? row._orderW : "—"}</td>}
                        <td style={{ padding: "8px 12px", color: "#3b82f6", fontWeight: 700, textAlign: "center" }}>{row._orderS > 0 ? row._orderS : "—"}</td>
                        <td style={{ padding: "8px 12px", color: "#a78bfa", textAlign: "right", whiteSpace: "nowrap" }}>{Number(row._totalPrice) > 0 ? row._totalPrice : "—"}</td>
                        <td style={{ padding: "8px 12px" }}>
                          <span style={{ padding: "3px 9px", borderRadius: 20, fontSize: 10, fontWeight: 600, color: decisionColor(d), background: decisionBg(d), whiteSpace: "nowrap" }}>{d}</span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {filtered.length > 500 && <div style={{ padding: 14, textAlign: "center", color: "#475569", fontSize: 11 }}>Showing 500 of {filtered.length} — Export CSV to see all</div>}
              {filtered.length === 0 && <div style={{ padding: 40, textAlign: "center", color: "#374151", fontSize: 13 }}>No products match the current filters</div>}
            </div>
          </div>
        )}

        {rawData && (
          <div style={{ marginTop: 10, display: "flex", justifyContent: "space-between", color: "#374151", fontSize: 11 }}>
            <span>Demand = Sales/Month × {period.value} | Desired = Demand − Stock {isSchulstrasse ? "Schulstrasse" : "Laufenburg"}</span>
            <span>{filtered.length.toLocaleString()} products shown</span>
          </div>
        )}
      </div>

      {/* Footer */}
      <div style={{ borderTop: "1px solid #1e2530", padding: "12px 28px", background: "#0a0d13", display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}>
        <span style={{ fontSize: 11, color: "#374151" }}>Developed by</span>
        <span style={{ fontSize: 11, fontWeight: 700, color: "#64748b", letterSpacing: "0.3px" }}>Dev Youness Louiza</span>
        <span style={{ fontSize: 11, color: "#1e2530" }}>•</span>
        <span style={{ fontSize: 11, color: "#374151" }}>C-Total Order Agent © {new Date().getFullYear()}</span>
      </div>
    </div>
  );
}