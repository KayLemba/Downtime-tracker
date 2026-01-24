import React from "react";
import { NavLink, Routes, Route, Navigate } from "react-router-dom";
import DowntimePage from "./DowntimePage";
import logo from "./assets/logo.png";

const linkStyle = ({ isActive }) => ({
  padding: "10px 12px",
  borderRadius: 12,
  textDecoration: "none",
  border: "1px solid rgba(255,255,255,0.22)",
  background: isActive ? "rgba(255,255,255,0.22)" : "rgba(255,255,255,0.10)",
  color: "#fff",
  fontWeight: 900,
});

export default function App() {
  return (
    <div className="container">
      <header className="header">
        <div className="brand">
          <img
            src={logo}
            alt="Company Logo"
            className="logo"
            onClick={() => (window.location.href = "/zamtel")}
            title="Go to Zamtel"
          />
          <div className="brandText">
            <div className="brandTitle">Downtime Tracker</div>
            <div className="brandSub">Call Center Monitoring</div>
          </div>
        </div>

        <nav className="nav">
          <NavLink to="/zamtel" style={linkStyle}>
            Zamtel Downtime
          </NavLink>
          <NavLink to="/mtn" style={linkStyle}>
            MTN Downtime
          </NavLink>
        </nav>
      </header>

      <Routes>
        <Route path="/" element={<Navigate to="/zamtel" replace />} />

        <Route
          path="/zamtel"
          element={
            <DowntimePage
              key="zamtel_page"
              pageName="Zamtel"
              title="Zamtel Downtime"
              storageKey="downtime_zamtel_ONLY"
              exportPrefix="zamtel_downtime"
            />
          }
        />

        <Route
          path="/mtn"
          element={
            <DowntimePage
              key="mtn_page"
              pageName="MTN"
              title="MTN Downtime"
              storageKey="downtime_mtn_ONLY"
              exportPrefix="mtn_downtime"
            />
          }
        />

        <Route path="*" element={<div className="card">Page not found</div>} />
      </Routes>
    </div>
  );
}
