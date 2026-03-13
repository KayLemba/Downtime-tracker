// src/DowntimePage.js
// Full recode with: isolated MTN/Zamtel storage, auto minutes (locked), validation,
// row Recalculate (green), Add Row (blue), Export (green), Clear All (red),
// and page themes via page-zamtel / page-mtn classes.

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const HEADERS = [
  "No:",
  "Date",
  "Start Time",
  "End Time",
  "Total Downtime (Minutes)",
  "Type (Internet/System)",
  "Department/Team Affected",
  "Reported By",
  "Reason/Notes",
  "Action Taken",
  "Resolved (Yes/No)",
  "Comments",
];

function safeParseJSON(value, fallback) {
  try {
    const parsed = JSON.parse(value);
    return parsed ?? fallback;
  } catch {
    return fallback;
  }
}

function blankRow(no) {
  return {
    "No:": no,
    "Date": "",
    "Start Time": "",
    "End Time": "",
    "Total Downtime (Minutes)": "",
    "Type (Internet/System)": "",
    "Department/Team Affected": "",
    "Reported By": "",
    "Reason/Notes": "",
    "Action Taken": "",
    "Resolved (Yes/No)": "",
    "Comments": "",
  };
}

function normalizeLoadedRows(loaded) {
  if (!Array.isArray(loaded) || loaded.length === 0) return [blankRow(1)];
  return loaded.map((r, idx) => ({ ...blankRow(idx + 1), ...r, "No:": idx + 1 }));
}

function parseMinutes(v) {
  const s = (v ?? "").toString().trim();
  if (!s) return 0;
  const n = Number(s);
  return Number.isFinite(n) && n > 0 ? n : 0;
}

function toISODateLocal(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function startOfWeekMonday(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const dow = x.getDay();
  const diff = dow === 0 ? -6 : 1 - dow;
  x.setDate(x.getDate() + diff);
  x.setHours(0, 0, 0, 0);
  return x;
}

function startOfMonth(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), 1);
  x.setHours(0, 0, 0, 0);
  return x;
}

function validateRow(r) {
  const dateStr = (r["Date"] ?? "").toString().trim();
  const start = (r["Start Time"] ?? "").toString().trim();
  const end = (r["End Time"] ?? "").toString().trim();

  if (!dateStr || !start || !end) return { ok: true, error: "" };
  if (start === end) return { ok: false, error: "End Time cannot equal Start Time." };
  return { ok: true, error: "" };
}

function computeMinutes(r) {
  const dateStr = (r["Date"] ?? "").toString().trim();
  const start = (r["Start Time"] ?? "").toString().trim();
  const end = (r["End Time"] ?? "").toString().trim();

  if (!dateStr || !start || !end) return "";

  const startDt = new Date(`${dateStr}T${start}:00`);
  let endDt = new Date(`${dateStr}T${end}:00`);

  if (Number.isNaN(startDt.getTime()) || Number.isNaN(endDt.getTime())) return "";

  // Cross-midnight support
  if (endDt.getTime() < startDt.getTime()) {
    endDt = new Date(endDt.getTime() + 24 * 60 * 60 * 1000);
  }

  const diffMin = Math.round((endDt.getTime() - startDt.getTime()) / 60000);
  return diffMin >= 0 ? diffMin : "";
}

function exportToXLSX(filename, headers, rows, sheetName) {
  const data = [headers, ...rows.map((r) => headers.map((h) => (r[h] ?? "").toString()))];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = headers.map((h) => ({
    wch: Math.min(46, Math.max(12, h.length + 2)),
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

function TotalsCard({ label, minutes }) {
  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;
  const display = hours > 0 ? `${hours}h ${mins}m` : `${mins}m`;

  return (
    <div className="card totalsCard">
      <div className="totalsLabel">{label}</div>
      <div className="totalsValue">{display}</div>
      <div className="totalsSub">{minutes} minutes</div>
    </div>
  );
}

function CellInput({ header, value, onChange, error }) {
  const v = value ?? "";
  const cls = `cell ${error ? "cellError" : ""}`;

  if (header === "Date") {
    return (
      <input
        className={cls}
        type="date"
        value={v}
        onChange={(e) => onChange(e.target.value)}
      />
    );
  }

  if (header === "Start Time" || header === "End Time") {
    return (
      <input
        className={cls}
        type="time"
        value={v}
        onChange={(e) => onChange(e.target.value)}
      />
    );
  }

  if (header === "Resolved (Yes/No)") {
    return (
      <select className={cls} value={v} onChange={(e) => onChange(e.target.value)}>
        <option value="">—</option>
        <option value="Yes">Yes</option>
        <option value="No">No</option>
      </select>
    );
  }

  if (header === "Type (Internet/System)") {
    return (
      <select className={cls} value={v} onChange={(e) => onChange(e.target.value)}>
        <option value="">—</option>
        <option value="Internet">Internet</option>
        <option value="System">System</option>
        <option value="SIP">SIP</option>
        <option value="Power">Power</option>
        <option value="Other">Other</option>
      </select>
    );
  }

  const long = header === "Reason/Notes" || header === "Action Taken" || header === "Comments";
  return (
    <input
      className={`${cls} ${long ? "cellLong" : ""}`}
      value={v}
      onChange={(e) => onChange(e.target.value)}
    />
  );
}

function LockedMinutesCell({ minutes, hasError }) {
  return (
    <div className={`lockedWrap ${hasError ? "lockedWrapError" : ""}`}>
      <span className="lockIcon" title="Auto-calculated (read-only)">
        🔒
      </span>
      <input className="cell lockedInput" value={minutes ?? ""} readOnly />
    </div>
  );
}

export default function DowntimePage({ pageName, title, storageKey, exportPrefix }) {
  const [rows, setRows] = useState([blankRow(1)]);
  const [search, setSearch] = useState("");

  useEffect(() => {
    const fromLS = localStorage.getItem(storageKey);
    const loaded = fromLS ? safeParseJSON(fromLS, []) : [];
    setRows(normalizeLoadedRows(loaded));
  }, [storageKey]);

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(rows));
  }, [rows, storageKey]);

  const validationMap = useMemo(() => {
    const map = {};
    for (const r of rows) map[r["No:"]] = validateRow(r);
    return map;
  }, [rows]);

  const totals = useMemo(() => {
    const now = new Date();
    const todayISO = toISODateLocal(now);
    const weekStart = startOfWeekMonday(now);
    const monthStart = startOfMonth(now);

    let today = 0;
    let week = 0;
    let month = 0;
    let all = 0;

    for (const r of rows) {
      const dateStr = (r["Date"] ?? "").toString().trim();
      const mins = parseMinutes(r["Total Downtime (Minutes)"]);
      if (!dateStr || mins <= 0) continue;

      all += mins;

      if (dateStr === todayISO) today += mins;

      const d = new Date(`${dateStr}T00:00:00`);
      if (!Number.isNaN(d.getTime())) {
        if (d >= weekStart && d <= now) week += mins;
        if (d >= monthStart && d <= now) month += mins;
      }
    }

    return { today, week, month, all };
  }, [rows]);

  const filteredRows = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return rows;
    return rows.filter((r) =>
      HEADERS.some((h) => (r[h] ?? "").toString().toLowerCase().includes(q))
    );
  }, [rows, search]);

  function renumber(list) {
    return list.map((r, idx) => ({ ...r, "No:": idx + 1 }));
  }

  function addRow() {
    setRows((prev) => renumber([...prev, blankRow(prev.length + 1)]));
  }

  function deleteRow(index) {
    setRows((prev) => {
      if (prev.length === 1) return [blankRow(1)];
      return renumber(prev.filter((_, i) => i !== index));
    });
  }

  function clearAll() {
    setRows([blankRow(1)]);
  }

  function recalcRow(index) {
    setRows((prev) =>
      prev.map((r, idx) => {
        if (idx !== index) return r;
        const v = validateRow(r);
        if (!v.ok) return { ...r, "Total Downtime (Minutes)": "" };
        return { ...r, "Total Downtime (Minutes)": computeMinutes(r) };
      })
    );
  }

  function updateCell(rowIndex, header, value) {
    setRows((prev) => {
      const next = prev.map((r, idx) => (idx === rowIndex ? { ...r, [header]: value } : r));

      if (header === "Date" || header === "Start Time" || header === "End Time") {
        const r = next[rowIndex];
        const v = validateRow(r);
        if (!v.ok) {
          next[rowIndex] = { ...r, "Total Downtime (Minutes)": "" };
        } else {
          next[rowIndex] = { ...r, "Total Downtime (Minutes)": computeMinutes(r) };
        }
      }

      return next;
    });
  }

  function exportXLSX() {
    const today = new Date().toISOString().slice(0, 10);
    exportToXLSX(`${exportPrefix}_${today}.xlsx`, HEADERS, rows, title);
  }

  const themeClass =
    (pageName ?? "").toString().toLowerCase() === "zamtel" ? "page-zamtel" : "page-mtn";

  return (
    <div className={`page ${themeClass}`}>
      <div className="banner">{`Currently recording: ${pageName}`}</div>

      <div className="pageTop">
        <h3 className="pageTitle">{title}</h3>

        <div className="toolbar">
          <input
            className="search"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search…"
          />

          <button className="btn btnPrimary" onClick={addRow}>
            + Add Row
          </button>

          <button className="btn btnSuccess" onClick={exportXLSX}>
            Export Excel (.xlsx)
          </button>

          <button className="btn btnDanger" onClick={clearAll}>
            Clear All
          </button>
        </div>
      </div>

      <div className="totalsGrid">
        <TotalsCard label="Today" minutes={totals.today} />
        <TotalsCard label="This Week" minutes={totals.week} />
        <TotalsCard label="This Month" minutes={totals.month} />
        <TotalsCard label="All Time" minutes={totals.all} />
      </div>

      <div className="tableWrap">
        <table className="table">
          <thead>
            <tr>
              {HEADERS.map((h) => (
                <th key={h}>{h}</th>
              ))}
              <th>Actions</th>
            </tr>
          </thead>

          <tbody>
            {filteredRows.map((row) => {
              const realIndex = rows.findIndex((r) => r["No:"] === row["No:"]);
              const v = validationMap[row["No:"]] || { ok: true, error: "" };
              const hasError = !v.ok;

              return (
                <tr key={row["No:"]}>
                  {HEADERS.map((h) => {
                    if (h === "No:") {
                      return (
                        <td key={h}>
                          <div className="noCell">{row[h]}</div>
                        </td>
                      );
                    }

                    if (h === "Total Downtime (Minutes)") {
                      return (
                        <td key={h}>
                          <LockedMinutesCell minutes={row[h]} hasError={hasError} />
                          {hasError ? <div className="errorText">{v.error}</div> : null}
                        </td>
                      );
                    }

                    const errorOnTime = hasError && (h === "Start Time" || h === "End Time");

                    return (
                      <td key={h}>
                        <CellInput
                          header={h}
                          value={row[h]}
                          error={errorOnTime}
                          onChange={(val) => updateCell(realIndex, h, val)}
                        />
                      </td>
                    );
                  })}

                  <td>
                    <div className="rowActions">
                      <button className="btn btnSuccess" onClick={() => recalcRow(realIndex)}>
                        Recalculate
                      </button>

                      <button className="btn btnDanger" onClick={() => deleteRow(realIndex)}>
                        Delete
                      </button>
                    </div>
                  </td>
                </tr>
              );
            })}

            {filteredRows.length === 0 ? (
              <tr>
                <td colSpan={HEADERS.length + 1} className="empty">
                  No matches.
                </td>
              </tr>
            ) : null}
          </tbody>
        </table>
      </div>
    </div>
  );
}