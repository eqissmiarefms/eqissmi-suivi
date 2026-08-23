import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  Plus, Trash2, Upload, Download, ChevronDown, ChevronRight,
  Users, CheckCircle2, Layers, FileSpreadsheet, Settings2,
  ClipboardList, LayoutDashboard, AlertCircle, Loader2, FileText, RefreshCw
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, PieChart, Pie, Cell,
} from "recharts";
import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType } from "docx";
import { saveAs } from "file-saver";
import { api } from "./api";
import logoAref from "./assets/logo-aref.png";

/* ---------- Palette (inspiration zellige : teal profond, terre cuite, sable) ---------- */
const C = {
  bg: "#F6F1E7",
  panel: "#FFFFFF",
  ink: "#20302C",
  inkSoft: "#5B6B63",
  teal: "#0E5C55",
  tealDark: "#0A423D",
  clay: "#BE6438",
  gold: "#B98A2B",
  good: "#3E7D5A",
  goodBg: "#E7F0E9",
  warn: "#B98A2B",
  warnBg: "#F5EEDD",
  bad: "#B24B32",
  badBg: "#F6E7E1",
  line: "#E4DBC7",
};

const POLL_INTERVAL_MS = 8000; // fréquence de synchronisation avec les autres utilisateurs

/* ---------- Motif géométrique (étoile à 8 branches, discret) ---------- */
function ZellijDivider() {
  const star = (x) => (
    <g key={x} transform={`translate(${x},0)`} opacity="0.55">
      <path
        d="M12 0 L15 9 L24 12 L15 15 L12 24 L9 15 L0 12 L9 9 Z"
        fill="none"
        stroke={C.gold}
        strokeWidth="1"
      />
    </g>
  );
  const xs = Array.from({ length: 14 }, (_, i) => i * 34);
  return (
    <svg width="100%" height="24" viewBox={`0 0 ${14 * 34} 24`} preserveAspectRatio="none" style={{ display: "block" }}>
      {xs.map((x) => star(x))}
    </svg>
  );
}

/* ---------- Logo AREF ----------
   PROVISOIRE : ceci est un emblème générique, pas le logo officiel de l'AREF
   (je n'ai pas le fichier). Pour utiliser le vrai logo :
   1. Placez le fichier (ex. logo-aref.png ou .svg) dans src/assets/
   2. En haut du fichier : import logoAref from "./assets/logo-aref.png";
   3. Remplacez <Logo /> ci-dessous par : <img src={logoAref} alt="AREF Marrakech-Safi" style={{ height: 46 }} />
------------------------------------------------------------------ */
function Logo() {
  return (
    <svg width="46" height="46" viewBox="0 0 46 46" aria-label="AREF Marrakech-Safi">
      <circle cx="23" cy="23" r="22" fill={C.tealDark} />
      <circle cx="23" cy="23" r="22" fill="none" stroke={C.gold} strokeWidth="1.2" />
      <path
        d="M23 8 L26.5 19.5 L38 23 L26.5 26.5 L23 38 L19.5 26.5 L8 23 L19.5 19.5 Z"
        fill="none" stroke={C.gold} strokeWidth="1.3"
      />
      <text x="23" y="26" textAnchor="middle" fontSize="9.5" fontWeight="700" fill="#fff" fontFamily="ui-serif, Georgia, serif">
        AREF
      </text>
    </svg>
  );
}

const norm = (s) => (s || "").toString().trim();
const normKey = (s) => norm(s).toLowerCase();

export default function EQissmiSuivi({ user, onLogout }) {
  const [sessions, setSessions] = useState([]);
  const [modules, setModules] = useState([]);
  const [entries, setEntries] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [syncState, setSyncState] = useState("idle"); // idle | syncing | synced | error
  const [tab, setTab] = useState("config");
  const fileInputRef = useRef(null);
  const [importReport, setImportReport] = useState(null);
  const pollingRef = useRef(null);

  const applyState = (data) => {
    setSessions(data.sessions || []);
    setModules(data.modules || []);
    setEntries(data.entries || []);
  };

  const refresh = async ({ silent } = {}) => {
    if (!silent) setSyncState("syncing");
    try {
      const data = await api.getState();
      applyState(data);
      setSyncState("synced");
    } catch (e) {
      setSyncState("error");
    }
  };

  /* ---------- Chargement initial + synchronisation périodique entre utilisateurs ---------- */
  useEffect(() => {
    (async () => {
      await refresh();
      setLoaded(true);
    })();
    pollingRef.current = setInterval(() => refresh({ silent: true }), POLL_INTERVAL_MS);
    return () => clearInterval(pollingRef.current);
  }, []);

  /* ---------- Helpers CRUD (chaque action écrit sur le serveur, puis resynchronise) ---------- */
  const addSession = async (name) => {
    const n = norm(name);
    if (!n) return;
    try {
      await api.createSession(n);
      await refresh({ silent: true });
    } catch (e) {
      setSyncState("error");
    }
  };
  const addModule = async (name) => {
    const n = norm(name);
    if (!n) return;
    try {
      await api.createModule(n);
      await refresh({ silent: true });
    } catch (e) {
      setSyncState("error");
    }
  };
  const deleteSession = async (id) => {
    setSessions((prev) => prev.filter((s) => s.id !== id)); // retrait optimiste
    try {
      await api.deleteSession(id);
      await refresh({ silent: true });
    } catch (e) {
      setSyncState("error");
      await refresh({ silent: true });
    }
  };
  const deleteModule = async (id) => {
    setModules((prev) => prev.filter((m) => m.id !== id));
    try {
      await api.deleteModule(id);
      await refresh({ silent: true });
    } catch (e) {
      setSyncState("error");
      await refresh({ silent: true });
    }
  };

  const addEntry = async (entry) => {
    try {
      await api.createEntry(entry);
      await refresh({ silent: true });
    } catch (e) {
      setSyncState("error");
    }
  };
  const deleteEntry = async (id) => {
    setEntries((prev) => prev.filter((e) => e.id !== id)); // retrait optimiste
    try {
      await api.deleteEntry(id);
      await refresh({ silent: true });
    } catch (e) {
      setSyncState("error");
      await refresh({ silent: true });
    }
  };
  const toggleSatisfied = async (id) => {
    const current = entries.find((e) => e.id === id);
    if (!current) return;
    const next = !current.satisfied;
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, satisfied: next } : e))); // optimiste
    try {
      await api.updateEntry(id, { satisfied: next });
      await refresh({ silent: true });
    } catch (e) {
      setSyncState("error");
      await refresh({ silent: true });
    }
  };

  const updateEntryProvince = async (id, province) => {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, province } : e))); // optimiste
    try {
      await api.updateEntry(id, { province });
      await refresh({ silent: true });
    } catch (e) {
      setSyncState("error");
      await refresh({ silent: true });
    }
  };

  /* ---------- Marquer "satisfait" pour des bénéficiaires déjà saisis, via une liste d'usernames ---------- */
  const markSatisfiedByUsernames = async (usernamesText) => {
    const wanted = usernamesText
      .split("\n")
      .map((u) => u.trim().toLowerCase())
      .filter(Boolean);
    const wantedSet = new Set(wanted);
    if (wantedSet.size === 0) return { updated: 0, alreadySatisfied: 0, notFound: [] };

    const matches = entries.filter((e) => e.username && wantedSet.has(e.username.trim().toLowerCase()));
    const toUpdate = matches.filter((e) => !e.satisfied);
    const alreadySatisfied = matches.length - toUpdate.length;
    const foundUsernames = new Set(matches.map((e) => e.username.trim().toLowerCase()));
    const notFound = wanted.filter((u) => !foundUsernames.has(u));

    if (toUpdate.length > 0) {
      const idsToUpdate = new Set(toUpdate.map((e) => e.id));
      setEntries((prev) => prev.map((e) => (idsToUpdate.has(e.id) ? { ...e, satisfied: true } : e))); // optimiste
      try {
        await Promise.all(toUpdate.map((e) => api.updateEntry(e.id, { satisfied: true })));
      } catch (err) {
        setSyncState("error");
      }
      await refresh({ silent: true });
    }

    return { updated: toUpdate.length, alreadySatisfied, notFound };
  };

  /* ---------- Agrégation pour le tableau de bord ---------- */
  const stats = useMemo(() => {
    const bySession = {};
    for (const s of sessions) {
      bySession[s.id] = { id: s.id, name: s.name, modules: {} };
    }
    for (const m of modules) {
      for (const s of sessions) {
        bySession[s.id].modules[m.id] = {
          id: m.id,
          name: m.name,
          tutors: {}, // tutorName -> { beneficiaries: Map(name->satisfied) }
        };
      }
    }
    for (const e of entries) {
      const sBucket = bySession[e.sessionId];
      if (!sBucket) continue;
      const mBucket = sBucket.modules[e.moduleId];
      if (!mBucket) continue;
      const tName = norm(e.tutor) || "(sans tuteur)";
      mBucket.tutors[tName] = mBucket.tutors[tName] || { beneficiaries: new Map() };
      const bName = norm(e.beneficiary) || "(sans nom)";
      mBucket.tutors[tName].beneficiaries.set(bName, !!e.satisfied);
    }

    // calcul des totaux par module
    const result = sessions.map((s) => {
      const sb = bySession[s.id];
      const modulesArr = modules
        .map((m) => {
          const mb = sb.modules[m.id];
          const tutorNames = Object.keys(mb.tutors);
          let total = 0;
          let satisfied = 0;
          const tutorRows = tutorNames.map((tn) => {
            const benef = mb.tutors[tn].beneficiaries;
            const count = benef.size;
            const sat = [...benef.values()].filter(Boolean).length;
            total += count;
            satisfied += sat;
            return { tutor: tn, count, sat };
          });
          const pct = total > 0 ? Math.round((satisfied / total) * 1000) / 10 : null;
          return {
            id: m.id,
            name: m.name,
            nbTuteurs: tutorNames.length,
            tutorRows,
            total,
            satisfied,
            pct,
          };
        })
        .filter((m) => m.total > 0 || m.nbTuteurs > 0);
      return { id: s.id, name: s.name, modules: modulesArr };
    });
    return result;
  }, [sessions, modules, entries]);

  /* ---------- Import Excel ---------- */
  const handleFile = async (file) => {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      const rows = rawRows.map((row) => {
        const keys = Object.keys(row).reduce((acc, k) => {
          acc[normKey(k)] = row[k];
          return acc;
        }, {});
        const statutRaw = normKey(keys["statut"] || keys["statut module"]);
        return {
          sessionName: norm(keys["session"]),
          moduleName: norm(keys["module"]),
          tutor: norm(keys["tuteur"]),
          beneficiary: norm(keys["beneficiaire"] || keys["bénéficiaire"]),
          username: norm(keys["username"] || keys["nom d'utilisateur"]),
          province: norm(keys["direction"] || keys["direction provinciale"] || keys["province"]),
          satisfied: ["oui", "satisfait", "1", "true", "vrai"].includes(statutRaw),
        };
      });

      const report = await api.importBulk(rows);
      setImportReport(report);
      await refresh({ silent: true });
    } catch (e) {
      setImportReport({ error: "Impossible de lire ce fichier, ou le serveur est injoignable. Vérifiez qu'il s'agit bien d'un .xlsx valide et que l'API tourne." });
    }
  };

  /* ---------- Export : modèle vide ---------- */
  const exportTemplate = () => {
    const data = [
      { Session: "Session 1 - 2026", Module: "Module 1 - Pédagogie numérique", Tuteur: "Nom du tuteur", "Bénéficiaire": "Nom du bénéficiaire", Username: "nom.utilisateur", "Direction provinciale": "Marrakech", Statut: "Oui" },
      { Session: "Session 1 - 2026", Module: "Module 1 - Pédagogie numérique", Tuteur: "Nom du tuteur", "Bénéficiaire": "Autre bénéficiaire", Username: "autre.utilisateur", "Direction provinciale": "Safi", Statut: "Non" },
    ];
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{ wch: 22 }, { wch: 28 }, { wch: 20 }, { wch: 24 }, { wch: 18 }, { wch: 18 }, { wch: 10 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Modèle");
    XLSX.writeFile(wb, "modele_import_eQissmi.xlsx");
  };

  /* ---------- Export : données complètes ---------- */
  const exportData = () => {
    const rows = entries.map((e) => ({
      Session: sessions.find((s) => s.id === e.sessionId)?.name || "",
      Module: modules.find((m) => m.id === e.moduleId)?.name || "",
      Tuteur: e.tutor,
      "Bénéficiaire": e.beneficiary,
      Username: e.username || "",
      "Direction provinciale": e.province || "",
      Statut: e.satisfied ? "Oui" : "Non",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 22 }, { wch: 28 }, { wch: 20 }, { wch: 24 }, { wch: 18 }, { wch: 18 }, { wch: 10 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Données");
    XLSX.writeFile(wb, "eQissmi_donnees.xlsx");
  };

  /* ---------- Export : tableau de bord agrégé ---------- */
  const exportDashboard = () => {
    const rows = [];
    stats.forEach((s) => {
      s.modules.forEach((m) => {
        rows.push({
          Session: s.name,
          Module: m.name,
          "Nb tuteurs": m.nbTuteurs,
          "Nb bénéficiaires": m.total,
          "Bénéficiaires satisfaits": m.satisfied,
          "% de satisfaction": m.pct === null ? "" : m.pct,
        });
      });
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 22 }, { wch: 28 }, { wch: 12 }, { wch: 16 }, { wch: 20 }, { wch: 16 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tableau de bord");
    XLSX.writeFile(wb, "eQissmi_tableau_de_bord.xlsx");
  };

  if (!loaded) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: 300, color: C.inkSoft, fontFamily: "ui-sans-serif, system-ui" }}>
        <Loader2 size={18} style={{ marginRight: 8, animation: "spin 1s linear infinite" }} />
        Chargement…
        <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
      </div>
    );
  }

  return (
    <div style={{ background: C.bg, color: C.ink, fontFamily: "ui-sans-serif, system-ui, sans-serif", minHeight: "100%", padding: "0" }}>
      <style>{`
        * { box-sizing: border-box; }
        input, select, textarea, button { font-family: inherit; }
        .eq-input {
          border: 1px solid ${C.line}; border-radius: 8px; padding: 8px 10px;
          font-size: 13.5px; background: #fff; color: ${C.ink}; width: 100%;
        }
        .eq-input:focus { outline: none; border-color: ${C.teal}; box-shadow: 0 0 0 3px rgba(14,92,85,0.12); }
        .eq-btn {
          display: inline-flex; align-items: center; gap: 6px; border-radius: 8px;
          padding: 8px 14px; font-size: 13.5px; font-weight: 600; cursor: pointer; border: none;
          transition: opacity .15s;
        }
        .eq-btn:hover { opacity: 0.88; }
        .eq-btn-primary { background: ${C.teal}; color: #fff; }
        .eq-btn-secondary { background: #fff; color: ${C.teal}; border: 1px solid ${C.teal}; }
        .eq-btn-ghost { background: transparent; color: ${C.inkSoft}; }
        .eq-btn-danger { background: transparent; color: ${C.bad}; }
        .eq-tab {
          display: flex; align-items: center; gap: 7px; padding: 10px 16px; font-size: 13.5px;
          font-weight: 600; cursor: pointer; border-bottom: 2px solid transparent; color: ${C.inkSoft};
        }
        .eq-tab.active { color: ${C.teal}; border-bottom-color: ${C.teal}; }
        .eq-card {
          background: ${C.panel}; border: 1px solid ${C.line}; border-radius: 12px; padding: 18px;
        }
        table.eq-table { width: 100%; border-collapse: collapse; font-size: 13px; }
        table.eq-table th { text-align: left; color: ${C.inkSoft}; font-weight: 600; font-size: 11.5px;
          text-transform: uppercase; letter-spacing: 0.04em; padding: 6px 10px; border-bottom: 1px solid ${C.line}; }
        table.eq-table td { padding: 8px 10px; border-bottom: 1px solid ${C.line}; }
      `}</style>

      {/* Header */}
      <div style={{ padding: "26px 28px 0" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div style={{ display: "flex", gap: 14, alignItems: "flex-start" }}>
            <img src={logoAref} alt="AREF Marrakech-Safi" style={{ height: 100 }} />
            <div>
              <div style={{ fontSize: 11.5, fontWeight: 700, letterSpacing: "0.08em", color: C.clay, textTransform: "uppercase" }}>
                e-Qissmi · AREF Marrakech-Safi
              </div>
              <h1 style={{ fontFamily: "ui-serif, Georgia, serif", fontSize: 26, margin: "4px 0 2px", color: C.tealDark }}>
                Suivi des sessions et modules
              </h1>
              <div style={{ color: C.inkSoft, fontSize: 13 }}>
                Tuteurs, bénéficiaires et taux de satisfaction par module
              </div>
            </div>
          </div>
          <div style={{ fontSize: 12, color: C.inkSoft, display: "flex", alignItems: "center", gap: 10, paddingTop: 4 }}>
            {syncState === "syncing" && <span style={{ display: "flex", alignItems: "center", gap: 6 }}><Loader2 size={13} style={{ animation: "spin 1s linear infinite" }} /> synchronisation…</span>}
            {syncState === "synced" && <span style={{ display: "flex", alignItems: "center", gap: 6 }}><CheckCircle2 size={13} color={C.good} /> synchronisé</span>}
            {syncState === "error" && <span style={{ display: "flex", alignItems: "center", gap: 6, color: C.bad }}><AlertCircle size={13} /> serveur injoignable</span>}
            <button className="eq-btn eq-btn-ghost" style={{ padding: "4px 8px" }} onClick={() => refresh()} title="Actualiser maintenant">
              Actualiser
            </button>
            {user && (
              <>
                <span style={{ color: C.line }}>|</span>
                <span>{user.email}</span>
                <button className="eq-btn eq-btn-ghost" style={{ padding: "4px 8px" }} onClick={onLogout}>
                  Déconnexion
                </button>
              </>
            )}
          </div>
        </div>
        <div style={{ margin: "16px 0 0" }}><ZellijDivider /></div>
      </div>

      {/* Tabs */}
      <div style={{ display: "flex", gap: 4, padding: "6px 22px 0", borderBottom: `1px solid ${C.line}` }}>
        <div className={`eq-tab ${tab === "config" ? "active" : ""}`} onClick={() => setTab("config")}>
          <Settings2 size={15} /> Configuration
        </div>
        <div className={`eq-tab ${tab === "saisie" ? "active" : ""}`} onClick={() => setTab("saisie")}>
          <ClipboardList size={15} /> Saisie
        </div>
        <div className={`eq-tab ${tab === "import" ? "active" : ""}`} onClick={() => setTab("import")}>
          <FileSpreadsheet size={15} /> Import / Export
        </div>
        <div className={`eq-tab ${tab === "dashboard" ? "active" : ""}`} onClick={() => setTab("dashboard")}>
          <LayoutDashboard size={15} /> Tableau de bord
        </div>
        <div className={`eq-tab ${tab === "rapport" ? "active" : ""}`} onClick={() => setTab("rapport")}>
          <FileText size={15} /> Rapport
        </div>
      </div>

      <div style={{ padding: 22 }}>
        {tab === "config" && <ConfigTab sessions={sessions} modules={modules} addSession={addSession} addModule={addModule} deleteSession={deleteSession} deleteModule={deleteModule} />}
        {tab === "saisie" && <SaisieTab sessions={sessions} modules={modules} entries={entries} addEntry={addEntry} deleteEntry={deleteEntry} toggleSatisfied={toggleSatisfied} updateEntryProvince={updateEntryProvince} markSatisfiedByUsernames={markSatisfiedByUsernames} />}
        {tab === "import" && (
          <ImportTab
            fileInputRef={fileInputRef}
            handleFile={handleFile}
            importReport={importReport}
            exportTemplate={exportTemplate}
            exportData={exportData}
            exportDashboard={exportDashboard}
            entriesCount={entries.length}
          />
        )}
        {tab === "dashboard" && <DashboardTab stats={stats} entries={entries} />}
        {tab === "rapport" && <ReportTab sessions={sessions} modules={modules} entries={entries} />}
      </div>
    </div>
  );
}

/* ============================================================= CONFIG ============================================================= */
function ConfigTab({ sessions, modules, addSession, addModule, deleteSession, deleteModule }) {
  const [sName, setSName] = useState("");
  const [mName, setMName] = useState("");
  return (
    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18 }}>
      <div className="eq-card">
        <h3 style={{ margin: "0 0 4px", fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>Sessions</h3>
        <p style={{ margin: "0 0 12px", fontSize: 12.5, color: C.inkSoft }}>Ex : "Session 1 - 2026", "Session printemps 2026"</p>
        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          <input className="eq-input" placeholder="Nom de la session" value={sName} onChange={(e) => setSName(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") { addSession(sName); setSName(""); } }} />
          <button className="eq-btn eq-btn-primary" onClick={() => { addSession(sName); setSName(""); }}><Plus size={14} /> Ajouter</button>
        </div>
        {sessions.length === 0 && <div style={{ fontSize: 12.5, color: C.inkSoft }}>Aucune session pour l'instant.</div>}
        {sessions.map((s) => (
          <div key={s.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${C.line}` }}>
            <span style={{ fontSize: 13.5 }}>{s.name}</span>
            <button className="eq-btn eq-btn-danger" onClick={() => deleteSession(s.id)}><Trash2 size={14} /></button>
          </div>
        ))}
      </div>

      <div className="eq-card">
        <h3 style={{ margin: "0 0 4px", fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>Modules</h3>
        <p style={{ margin: "0 0 12px", fontSize: 12.5, color: C.inkSoft }}>Catalogue commun à toutes les sessions</p>
        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          <input className="eq-input" placeholder="Nom du module" value={mName} onChange={(e) => setMName(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") { addModule(mName); setMName(""); } }} />
          <button className="eq-btn eq-btn-primary" onClick={() => { addModule(mName); setMName(""); }}><Plus size={14} /> Ajouter</button>
        </div>
        {modules.length === 0 && <div style={{ fontSize: 12.5, color: C.inkSoft }}>Aucun module pour l'instant.</div>}
        {modules.map((m) => (
          <div key={m.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${C.line}` }}>
            <span style={{ fontSize: 13.5 }}>{m.name}</span>
            <button className="eq-btn eq-btn-danger" onClick={() => deleteModule(m.id)}><Trash2 size={14} /></button>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ============================================================= SAISIE ============================================================= */
const DIRECTIONS = ["Marrakech", "Safi", "Rhamna", "Chichaoua", "Kalaa des Sraghna", "Youssoufia", "Haouz", "Essaouira"];
const DIRECTIONS_AR = {
  "Marrakech": "مراكش",
  "Safi": "آسفي",
  "Rhamna": "الرحامنة",
  "Chichaoua": "شيشاوة",
  "Kalaa des Sraghna": "قلعة السراغنة",
  "Youssoufia": "اليوسفية",
  "Haouz": "الحوز",
  "Essaouira": "الصويرة",
  "Non renseigné": "غير محدد",
};

function emptyRow() {
  return { key: Math.random().toString(36).slice(2), name: "", username: "", province: "" };
}

function SaisieTab({ sessions, modules, entries, addEntry, deleteEntry, toggleSatisfied, updateEntryProvince, markSatisfiedByUsernames }) {
  const [sessionId, setSessionId] = useState("");
  const [moduleId, setModuleId] = useState("");
  const [tutor, setTutor] = useState("");
  const [rows, setRows] = useState([emptyRow()]);
  const [satisfiedUsernamesText, setSatisfiedUsernamesText] = useState("");
  const [filterSession, setFilterSession] = useState("");
  const [filterModule, setFilterModule] = useState("");
  const [filterDirection, setFilterDirection] = useState("");
  const [searchText, setSearchText] = useState("");

  const [updateUsernamesText, setUpdateUsernamesText] = useState("");
  const [updateReport, setUpdateReport] = useState(null);
  const [updating, setUpdating] = useState(false);

  const submitUpdate = async () => {
    if (!updateUsernamesText.trim()) return;
    setUpdating(true);
    const report = await markSatisfiedByUsernames(updateUsernamesText);
    setUpdateReport(report);
    setUpdating(false);
    setUpdateUsernamesText("");
  };

  const updateRow = (key, field, value) => {
    setRows((prev) => prev.map((r) => (r.key === key ? { ...r, [field]: value } : r)));
  };
  const addRow = () => setRows((prev) => [...prev, emptyRow()]);
  const removeRow = (key) => setRows((prev) => (prev.length > 1 ? prev.filter((r) => r.key !== key) : prev));

  const submit = () => {
    const validRows = rows.filter((r) => r.name.trim());
    if (!sessionId || !moduleId || !tutor.trim() || validRows.length === 0) return;
    const satisfiedUsernames = new Set(
      satisfiedUsernamesText.split("\n").map((u) => u.trim().toLowerCase()).filter(Boolean)
    );
    validRows.forEach((r) => {
      const name = r.name.trim();
      const username = r.username.trim();
      const satisfied = username ? satisfiedUsernames.has(username.toLowerCase()) : false;
      addEntry({ sessionId, moduleId, tutor: tutor.trim(), beneficiary: name, username, province: r.province, satisfied });
    });
    setRows([emptyRow()]);
    setSatisfiedUsernamesText("");
  };

  const filtered = entries.filter((e) => {
    const search = searchText.trim().toLowerCase();
    const matchesSearch =
      !search ||
      (e.beneficiary || "").toLowerCase().includes(search) ||
      (e.username || "").toLowerCase().includes(search);
    return (
      (!filterSession || e.sessionId === filterSession) &&
      (!filterModule || e.moduleId === filterModule) &&
      (!filterDirection || e.province === filterDirection) &&
      matchesSearch
    );
  });

  return (
    <div style={{ display: "grid", gridTemplateColumns: "440px 1fr", gap: 18 }}>
      <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <div className="eq-card">
        <h3 style={{ margin: "0 0 4px", fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>Ajout rapide</h3>
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.inkSoft }}>
          Un tuteur, plusieurs bénéficiaires — direction provinciale par bénéficiaire
        </p>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <select className="eq-input" value={sessionId} onChange={(e) => setSessionId(e.target.value)}>
            <option value="">Session…</option>
            {sessions.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
          <select className="eq-input" value={moduleId} onChange={(e) => setModuleId(e.target.value)}>
            <option value="">Module…</option>
            {modules.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}
          </select>
          <input className="eq-input" placeholder="Nom du tuteur" value={tutor} onChange={(e) => setTutor(e.target.value)} />

          <div style={{ fontSize: 12.5, color: C.inkSoft, marginTop: 4 }}>Bénéficiaires</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {rows.map((r) => (
              <div key={r.key} style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr auto", gap: 6, alignItems: "center" }}>
                <input
                  className="eq-input" placeholder="Nom du bénéficiaire" value={r.name}
                  onChange={(e) => updateRow(r.key, "name", e.target.value)}
                />
                <input
                  className="eq-input" placeholder="Username" value={r.username}
                  onChange={(e) => updateRow(r.key, "username", e.target.value)}
                />
                <select
                  className="eq-input" value={r.province}
                  onChange={(e) => updateRow(r.key, "province", e.target.value)}
                >
                  <option value="">Direction…</option>
                  {DIRECTIONS.map((d) => <option key={d} value={d}>{d}</option>)}
                </select>
                <button
                  className="eq-btn eq-btn-ghost" style={{ padding: "6px" }}
                  onClick={() => removeRow(r.key)} title="Retirer cette ligne"
                >
                  <Trash2 size={14} />
                </button>
              </div>
            ))}
          </div>
          <button className="eq-btn eq-btn-secondary" onClick={addRow} style={{ justifyContent: "center" }}>
            <Plus size={14} /> Ajouter une ligne
          </button>

          <div>
            <label style={{ display: "block", fontSize: 12.5, color: C.inkSoft, marginBottom: 6 }}>
              Usernames des bénéficiaires ayant satisfait le module (un par ligne)
            </label>
            <textarea
              className="eq-input"
              rows={4}
              placeholder={"ahmed.elfassi\nsalma.idrissi"}
              value={satisfiedUsernamesText}
              onChange={(e) => setSatisfiedUsernamesText(e.target.value)}
            />
            <div style={{ fontSize: 11.5, color: C.inkSoft, marginTop: 4 }}>
              Seuls les bénéficiaires dont le username figure ici seront marqués comme ayant satisfait.
            </div>
          </div>
          <button className="eq-btn eq-btn-primary" onClick={submit} style={{ justifyContent: "center" }}>
            <Plus size={14} /> Ajouter au suivi
          </button>
        </div>
        {(sessions.length === 0 || modules.length === 0) && (
          <div style={{ marginTop: 12, fontSize: 12, color: C.clay, display: "flex", gap: 6 }}>
            <AlertCircle size={14} style={{ flexShrink: 0, marginTop: 1 }} />
            Créez d'abord au moins une session et un module dans l'onglet Configuration.
          </div>
        )}
      </div>

      <div className="eq-card">
        <h3 style={{ margin: "0 0 4px", fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>
          Mettre à jour un statut existant
        </h3>
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.inkSoft }}>
          Pour des bénéficiaires déjà saisis : collez ici les usernames de ceux qui viennent de satisfaire le module.
          Les autres restent inchangés.
        </p>
        <textarea
          className="eq-input"
          rows={6}
          placeholder={"ahmed.elfassi\nsalma.idrissi\n…"}
          value={updateUsernamesText}
          onChange={(e) => setUpdateUsernamesText(e.target.value)}
        />
        <button
          className="eq-btn eq-btn-primary"
          onClick={submitUpdate}
          disabled={updating}
          style={{ justifyContent: "center", width: "100%", marginTop: 10 }}
        >
          <CheckCircle2 size={14} /> {updating ? "Mise à jour…" : "Marquer comme satisfait"}
        </button>
        {updateReport && (
          <div style={{ marginTop: 12, fontSize: 12.5, background: C.goodBg, color: C.good, borderRadius: 8, padding: 10 }}>
            {updateReport.updated} bénéficiaire(s) mis à jour
            {updateReport.alreadySatisfied > 0 && `, ${updateReport.alreadySatisfied} déjà marqué(s) satisfait`}
            {updateReport.notFound.length > 0 && (
              <div style={{ marginTop: 6, color: C.bad }}>
                Username(s) introuvable(s) : {updateReport.notFound.join(", ")}
              </div>
            )}
          </div>
        )}
      </div>
      </div>

      <div className="eq-card">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
          <h3 style={{ margin: 0, fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>
            Bénéficiaires saisis ({filtered.length})
          </h3>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <input
              className="eq-input"
              style={{ width: 200 }}
              placeholder="Rechercher (nom ou username)…"
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
            />
            <select className="eq-input" style={{ width: 150 }} value={filterSession} onChange={(e) => setFilterSession(e.target.value)}>
              <option value="">Toutes sessions</option>
              {sessions.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            <select className="eq-input" style={{ width: 160 }} value={filterModule} onChange={(e) => setFilterModule(e.target.value)}>
              <option value="">Tous modules</option>
              {modules.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}
            </select>
            <select className="eq-input" style={{ width: 150 }} value={filterDirection} onChange={(e) => setFilterDirection(e.target.value)}>
              <option value="">Toutes directions</option>
              {DIRECTIONS.map((d) => <option key={d} value={d}>{d}</option>)}
            </select>
          </div>
        </div>
        <div style={{ maxHeight: 480, overflowY: "auto" }}>
          <table className="eq-table">
            <thead>
              <tr>
                <th>Session</th><th>Module</th><th>Tuteur</th><th>Bénéficiaire</th><th>Direction</th><th>Username</th><th>Satisfait</th><th></th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((e) => (
                <tr key={e.id}>
                  <td>{sessions.find((s) => s.id === e.sessionId)?.name || "—"}</td>
                  <td>{modules.find((m) => m.id === e.moduleId)?.name || "—"}</td>
                  <td>{e.tutor}</td>
                  <td>{e.beneficiary}</td>
                  <td>
                    <select
                      className="eq-input"
                      style={{ padding: "4px 6px", fontSize: 12.5 }}
                      value={e.province || ""}
                      onChange={(ev) => updateEntryProvince(e.id, ev.target.value)}
                    >
                      <option value="">—</option>
                      {DIRECTIONS.map((d) => <option key={d} value={d}>{d}</option>)}
                    </select>
                  </td>
                  <td>{e.username || "—"}</td>
                  <td><input type="checkbox" checked={!!e.satisfied} onChange={() => toggleSatisfied(e.id)} /></td>
                  <td><button className="eq-btn eq-btn-ghost" onClick={() => deleteEntry(e.id)}><Trash2 size={13} /></button></td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr><td colSpan={8} style={{ color: C.inkSoft, textAlign: "center", padding: 20 }}>Aucune donnée pour ce filtre.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ============================================================= IMPORT / EXPORT ============================================================= */
function ImportTab({ fileInputRef, handleFile, importReport, exportTemplate, exportData, exportDashboard, entriesCount }) {
  const [dragOver, setDragOver] = useState(false);
  return (
    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18 }}>
      <div className="eq-card">
        <h3 style={{ margin: "0 0 4px", fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>Importer un fichier Excel</h3>
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.inkSoft }}>
          Colonnes attendues : <b>Session, Module, Tuteur, Bénéficiaire, Username, Statut</b> (Statut = "Oui" / "Non", Username optionnel).
          Les sessions et modules absents sont créés automatiquement.
        </p>
        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); setDragOver(false); const f = e.dataTransfer.files?.[0]; if (f) handleFile(f); }}
          onClick={() => fileInputRef.current?.click()}
          style={{
            border: `2px dashed ${dragOver ? C.teal : C.line}`, borderRadius: 12, padding: 30,
            textAlign: "center", cursor: "pointer", background: dragOver ? "#F0F6F4" : "#FAF7EF",
          }}
        >
          <Upload size={22} color={C.teal} style={{ marginBottom: 8 }} />
          <div style={{ fontSize: 13.5, fontWeight: 600, color: C.tealDark }}>Cliquez ou déposez un fichier .xlsx</div>
          <div style={{ fontSize: 12, color: C.inkSoft, marginTop: 2 }}>Les lignes valides sont ajoutées aux données existantes</div>
        </div>
        <input
          ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
          onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f); e.target.value = ""; }}
        />
        {importReport && (
          <div style={{
            marginTop: 14, padding: 12, borderRadius: 8, fontSize: 13,
            background: importReport.error ? C.badBg : C.goodBg, color: importReport.error ? C.bad : C.good,
          }}>
            {importReport.error
              ? importReport.error
              : `${importReport.added} ligne(s) importée(s)${importReport.skipped ? `, ${importReport.skipped} ignorée(s) (données manquantes)` : ""}.`}
          </div>
        )}
        <button className="eq-btn eq-btn-secondary" style={{ marginTop: 14 }} onClick={exportTemplate}>
          <Download size={14} /> Télécharger le modèle vide
        </button>
      </div>

      <div className="eq-card">
        <h3 style={{ margin: "0 0 4px", fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>Exporter</h3>
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.inkSoft }}>{entriesCount} bénéficiaire(s) actuellement enregistré(s)</p>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <button className="eq-btn eq-btn-primary" onClick={exportData}><Download size={14} /> Exporter toutes les données (détail)</button>
          <button className="eq-btn eq-btn-primary" onClick={exportDashboard}><Download size={14} /> Exporter le tableau de bord (agrégé)</button>
        </div>
        <div style={{ marginTop: 18, paddingTop: 14, borderTop: `1px solid ${C.line}`, fontSize: 12.5, color: C.inkSoft }}>
          <b style={{ color: C.ink }}>Conseil :</b> pour un grand volume de bénéficiaires, préparez le fichier en Excel puis importez-le
          ici plutôt que de saisir un par un. Vous pouvez importer plusieurs fichiers successivement, les données s'additionnent.
        </div>
      </div>
    </div>
  );
}

/* ============================================================= DASHBOARD ============================================================= */
function pctColor(pct) {
  if (pct === null) return { fg: C.inkSoft, bg: "#EFEBE0" };
  if (pct >= 70) return { fg: C.good, bg: C.goodBg };
  if (pct >= 40) return { fg: C.warn, bg: C.warnBg };
  return { fg: C.bad, bg: C.badBg };
}

function DashboardTab({ stats, entries }) {
  const [openModules, setOpenModules] = useState({});
  const toggle = (key) => setOpenModules((p) => ({ ...p, [key]: !p[key] }));
  const [filterSessionId, setFilterSessionId] = useState("");

  const activeSessions = stats.filter((s) => s.modules.length > 0);
  const hasData = activeSessions.length > 0;
  const visibleSessions = filterSessionId ? activeSessions.filter((s) => s.id === filterSessionId) : activeSessions;

  const provinceStatsBySession = useMemo(() => {
    const order = [...DIRECTIONS, "Non renseigné"];
    const result = {};
    activeSessions.forEach((s) => {
      const buckets = {};
      (entries || [])
        .filter((e) => e.sessionId === s.id)
        .forEach((e) => {
          const key = e.province && e.province.trim() ? e.province.trim() : "Non renseigné";
          buckets[key] = buckets[key] || { total: 0, satisfied: 0 };
          buckets[key].total += 1;
          if (e.satisfied) buckets[key].satisfied += 1;
        });
      result[s.id] = order
        .filter((name) => buckets[name])
        .map((name) => ({
          name,
          Bénéficiaires: buckets[name].total,
          "Ont satisfait": buckets[name].satisfied,
          pct: buckets[name].total > 0 ? Math.round((buckets[name].satisfied / buckets[name].total) * 1000) / 10 : null,
        }));
    });
    return result;
  }, [entries, stats]);

  if (!hasData) {
    return (
      <div className="eq-card" style={{ textAlign: "center", padding: 40, color: C.inkSoft }}>
        <Layers size={22} style={{ marginBottom: 8, opacity: 0.6 }} />
        <div>Aucune donnée pour le moment. Saisissez ou importez des bénéficiaires pour voir le tableau de bord.</div>
      </div>
    );
  }

  // Aperçu global : taux de satisfaction moyen par session
  const overview = activeSessions.map((s) => {
    const total = s.modules.reduce((a, m) => a + m.total, 0);
    const satisfied = s.modules.reduce((a, m) => a + m.satisfied, 0);
    return {
      name: s.name.length > 16 ? s.name.slice(0, 15) + "…" : s.name,
      Bénéficiaires: total,
      "Ont satisfait": satisfied,
    };
  });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <div className="eq-card" style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px 18px" }}>
        <span style={{ fontSize: 12.5, color: C.inkSoft, fontWeight: 600 }}>Filtrer par session</span>
        <select className="eq-input" style={{ width: 220 }} value={filterSessionId} onChange={(e) => setFilterSessionId(e.target.value)}>
          <option value="">Toutes les sessions</option>
          {activeSessions.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
        </select>
      </div>

      <div style={{ fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 17, margin: "4px 0 -6px" }}>
        Par direction provinciale
      </div>
      {visibleSessions.map((s) => {
        const data = provinceStatsBySession[s.id] || [];
        if (data.length === 0) return null;
        return (
          <div key={"province-" + s.id} className="eq-card">
            <h3 style={{ margin: "0 0 4px", fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>
              {s.name}
            </h3>
            <p style={{ margin: "0 0 10px", fontSize: 12.5, color: C.inkSoft }}>Bénéficiaires vs. bénéficiaires ayant satisfait, par direction provinciale</p>
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={data} margin={{ top: 4, right: 12, left: -12, bottom: 24 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 11, fill: C.inkSoft }} interval={0} angle={-20} textAnchor="end" height={50} />
                <YAxis tick={{ fontSize: 11.5, fill: C.inkSoft }} allowDecimals={false} />
                <Tooltip contentStyle={{ fontSize: 12.5, borderRadius: 8, border: `1px solid ${C.line}` }} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                <Bar dataKey="Bénéficiaires" fill={C.gold} radius={[4, 4, 0, 0]} />
                <Bar dataKey="Ont satisfait" fill={C.teal} radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
            <table className="eq-table" style={{ marginTop: 14 }}>
              <thead>
                <tr>
                  <th>Direction</th>
                  <th>Bénéficiaires</th>
                  <th>Ont satisfait</th>
                  <th>% satisfaction</th>
                </tr>
              </thead>
              <tbody>
                {data.map((p) => {
                  const col = pctColor(p.pct);
                  return (
                    <tr key={p.name}>
                      <td style={{ fontWeight: 600 }}>{p.name}</td>
                      <td>{p.Bénéficiaires}</td>
                      <td>{p["Ont satisfait"]}</td>
                      <td>
                        <span style={{ background: col.bg, color: col.fg, padding: "3px 9px", borderRadius: 999, fontWeight: 700, fontSize: 12 }}>
                          {p.pct === null ? "—" : `${p.pct}%`}
                        </span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        );
      })}

      {!filterSessionId && activeSessions.length > 1 && (
        <div className="eq-card">
          <h3 style={{ margin: "0 0 4px", fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>
            Aperçu global — toutes sessions
          </h3>
          <p style={{ margin: "0 0 10px", fontSize: 12.5, color: C.inkSoft }}>Bénéficiaires vs. bénéficiaires ayant satisfait, par session</p>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={overview} margin={{ top: 4, right: 12, left: -12, bottom: 4 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
              <XAxis dataKey="name" tick={{ fontSize: 11.5, fill: C.inkSoft }} />
              <YAxis tick={{ fontSize: 11.5, fill: C.inkSoft }} allowDecimals={false} />
              <Tooltip contentStyle={{ fontSize: 12.5, borderRadius: 8, border: `1px solid ${C.line}` }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Bénéficiaires" fill={C.gold} radius={[4, 4, 0, 0]} />
              <Bar dataKey="Ont satisfait" fill={C.teal} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      )}

      {visibleSessions.map((s) => {
        const total = s.modules.reduce((a, m) => a + m.total, 0);
        const satisfied = s.modules.reduce((a, m) => a + m.satisfied, 0);
        const donutData = [
          { name: "Ont satisfait", value: satisfied },
          { name: "N'ont pas satisfait", value: Math.max(total - satisfied, 0) },
        ];
        const moduleChartData = s.modules.map((m) => ({
          name: m.name.length > 14 ? m.name.slice(0, 13) + "…" : m.name,
          Bénéficiaires: m.total,
          "Ont satisfait": m.satisfied,
        }));

        return (
        <div key={s.id} className="eq-card">
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
            <div style={{ width: 8, height: 8, borderRadius: 4, background: C.clay }} />
            <h3 style={{ margin: 0, fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 17 }}>{s.name}</h3>
          </div>
          <div style={{ margin: "10px 0 12px" }}><ZellijDivider /></div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 220px", gap: 16, marginBottom: 18 }}>
            <div>
              <div style={{ fontSize: 12, fontWeight: 600, color: C.inkSoft, marginBottom: 6 }}>Par module</div>
              <ResponsiveContainer width="100%" height={200}>
                <BarChart data={moduleChartData} margin={{ top: 4, right: 12, left: -12, bottom: 4 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
                  <XAxis dataKey="name" tick={{ fontSize: 11, fill: C.inkSoft }} interval={0} angle={-12} textAnchor="end" height={44} />
                  <YAxis tick={{ fontSize: 11, fill: C.inkSoft }} allowDecimals={false} />
                  <Tooltip contentStyle={{ fontSize: 12.5, borderRadius: 8, border: `1px solid ${C.line}` }} />
                  <Bar dataKey="Bénéficiaires" fill={C.gold} radius={[4, 4, 0, 0]} />
                  <Bar dataKey="Ont satisfait" fill={C.teal} radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div>
              <div style={{ fontSize: 12, fontWeight: 600, color: C.inkSoft, marginBottom: 6, textAlign: "center" }}>
                Taux global — {total > 0 ? Math.round((satisfied / total) * 1000) / 10 : 0}%
              </div>
              <ResponsiveContainer width="100%" height={190}>
                <PieChart>
                  <Pie data={donutData} dataKey="value" innerRadius={45} outerRadius={70} paddingAngle={2}>
                    <Cell fill={C.teal} />
                    <Cell fill={C.line} />
                  </Pie>
                  <Tooltip contentStyle={{ fontSize: 12.5, borderRadius: 8, border: `1px solid ${C.line}` }} />
                </PieChart>
              </ResponsiveContainer>
            </div>
          </div>

          <table className="eq-table">
            <thead>
              <tr>
                <th style={{ width: 30 }}></th>
                <th>Module</th>
                <th><Users size={12} style={{ verticalAlign: -1 }} /> Tuteurs</th>
                <th>Bénéficiaires</th>
                <th>Ont satisfait</th>
                <th>% satisfaction</th>
              </tr>
            </thead>
            <tbody>
              {s.modules.map((m) => {
                const key = s.id + "-" + m.id;
                const open = !!openModules[key];
                const col = pctColor(m.pct);
                return (
                  <React.Fragment key={m.id}>
                    <tr style={{ cursor: "pointer" }} onClick={() => toggle(key)}>
                      <td>{open ? <ChevronDown size={14} /> : <ChevronRight size={14} />}</td>
                      <td style={{ fontWeight: 600 }}>{m.name}</td>
                      <td>{m.nbTuteurs}</td>
                      <td>{m.total}</td>
                      <td>{m.satisfied}</td>
                      <td>
                        <span style={{ background: col.bg, color: col.fg, padding: "3px 9px", borderRadius: 999, fontWeight: 700, fontSize: 12 }}>
                          {m.pct === null ? "—" : `${m.pct}%`}
                        </span>
                      </td>
                    </tr>
                    {open && (
                      <tr>
                        <td></td>
                        <td colSpan={5} style={{ padding: "4px 10px 12px" }}>
                          <table className="eq-table" style={{ background: "#FAF7EF", borderRadius: 8 }}>
                            <thead>
                              <tr><th>Tuteur</th><th>Nb bénéficiaires</th><th>Ont satisfait</th></tr>
                            </thead>
                            <tbody>
                              {m.tutorRows.map((t) => (
                                <tr key={t.tutor}>
                                  <td>{t.tutor}</td>
                                  <td>{t.count}</td>
                                  <td>{t.sat} / {t.count}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
        );
      })}
    </div>
  );
}

/* ============================================================= RAPPORT ============================================================= */
function computeReportStats(entries, sessionId, moduleId) {
  // Inscrits = bénéficiaires distincts dans la session (tous modules confondus), regroupés par direction
  const seenKeyToProvince = new Map();
  entries.forEach((e) => {
    if (e.sessionId !== sessionId) return;
    const key = e.username && e.username.trim() ? "u:" + e.username.trim().toLowerCase() : "n:" + (e.beneficiary || "").trim().toLowerCase();
    if (!seenKeyToProvince.has(key)) {
      seenKeyToProvince.set(key, e.province && e.province.trim() ? e.province.trim() : "Non renseigné");
    }
  });
  const registeredByProvince = {};
  seenKeyToProvince.forEach((province) => {
    registeredByProvince[province] = (registeredByProvince[province] || 0) + 1;
  });

  // Ont satisfait = pour le module choisi précisément, par direction
  const completedByProvince = {};
  entries.forEach((e) => {
    if (e.sessionId !== sessionId || e.moduleId !== moduleId || !e.satisfied) return;
    const province = e.province && e.province.trim() ? e.province.trim() : "Non renseigné";
    completedByProvince[province] = (completedByProvince[province] || 0) + 1;
  });

  const provinces = [...DIRECTIONS, "Non renseigné"].filter((p) => registeredByProvince[p] || completedByProvince[p]);
  const rows = provinces.map((p) => {
    const registered = registeredByProvince[p] || 0;
    const completed = completedByProvince[p] || 0;
    const pct = registered > 0 ? Math.round((completed / registered) * 1000) / 10 : null;
    return { province: p, registered, completed, pct };
  });

  const totalRegistered = rows.reduce((a, r) => a + r.registered, 0);
  const totalCompleted = rows.reduce((a, r) => a + r.completed, 0);
  const globalPct = totalRegistered > 0 ? Math.round((totalCompleted / totalRegistered) * 1000) / 10 : null;

  // Équipe régionale = tuteurs distincts de la session, par direction
  const teamMap = new Map();
  entries.forEach((e) => {
    if (e.sessionId !== sessionId) return;
    const tutor = (e.tutor || "").trim();
    if (!tutor) return;
    const province = e.province && e.province.trim() ? e.province.trim() : "Non renseigné";
    teamMap.set(province + "|||" + tutor.toLowerCase(), { province, tutor });
  });
  const team = [...teamMap.values()].sort((a, b) => a.province.localeCompare(b.province) || a.tutor.localeCompare(b.tutor));

  return { rows, totalRegistered, totalCompleted, globalPct, team };
}

function defaultSummary(sessionName, moduleName, stats) {
  return `Dans le cadre du suivi de la session "${sessionName}", un total de ${stats.totalRegistered} bénéficiaire(s) a été enregistré, réparti(s) sur ${stats.rows.length} direction(s) provinciale(s). Concernant le module "${moduleName}", ${stats.totalCompleted} bénéficiaire(s) ont satisfait aux exigences, soit un taux de réussite global de ${stats.globalPct !== null ? stats.globalPct + "%" : "non disponible"}. Ce suivi a été assuré par ${stats.team.length} tuteur(s) mobilisé(s) au niveau régional.`;
}

function defaultConclusion(sessionName, moduleName) {
  return `Les résultats obtenus pour le module "${moduleName}" de la session "${sessionName}" témoignent d'une dynamique positive dans l'accompagnement des bénéficiaires. Il convient de poursuivre les efforts d'encadrement technique et pédagogique afin de consolider ces acquis et d'améliorer davantage le taux de satisfaction dans les prochaines étapes de la formation.`;
}

function defaultSummaryAr(sessionName, moduleName, stats) {
  return `في إطار تتبع سير الدورة التكوينية "${sessionName}"، تم تسجيل ما مجموعه ${stats.totalRegistered} مستفيدا موزعين على ${stats.rows.length} مديرية إقليمية. أما بخصوص الوحدة "${moduleName}"، فقد تمكن ${stats.totalCompleted} مستفيدا من إنجازها، أي بنسبة إنجاز إجمالية بلغت ${stats.globalPct !== null ? stats.globalPct + "%" : "غير متوفرة"}. وقد تم تأطير هذا المسار من طرف ${stats.team.length} ميسرا على المستوى الجهوي.`;
}

function defaultConclusionAr(sessionName, moduleName) {
  return `تعكس النتائج المحصل عليها بخصوص الوحدة "${moduleName}" من الدورة "${sessionName}" دينامية إيجابية في مواكبة المستفيدين والمستفيدات. وتجدر الإشارة إلى ضرورة الاستمرار في تعزيز التأطير التقني والبيداغوجي من أجل ترسيخ هذه المكتسبات والرفع من نسب الإنجاز في المراحل المقبلة من التكوين.`;
}

async function generateReportDocxArabic({ orgLine1, orgLine2, title, sessionName, moduleName, preparedBy, coordinatedWith, summary, conclusion, stats }) {
  const P = (text, opts = {}) =>
    new Paragraph({
      bidirectional: true,
      alignment: opts.alignment || AlignmentType.RIGHT,
      heading: opts.heading,
      spacing: opts.spacing,
      children: [new TextRun({ text: String(text), bold: !!opts.bold, rightToLeft: true, font: "Arial" })],
    });

  const cellText = (text, opts = {}) =>
    new TableCell({
      children: [P(text, { bold: opts.bold })],
      width: { size: opts.width || 25, type: WidthType.PERCENTAGE },
    });

  const provName = (p) => DIRECTIONS_AR[p] || p;

  const registrationTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [cellText("المديرية الإقليمية", { bold: true }), cellText("عدد المسجلين", { bold: true }), cellText("عدد من أنجزوا الوحدة", { bold: true }), cellText("نسبة الإنجاز", { bold: true })] }),
      ...stats.rows.map((r) => new TableRow({ children: [cellText(provName(r.province)), cellText(r.registered), cellText(r.completed), cellText(r.pct === null ? "—" : r.pct + "%")] })),
      new TableRow({ children: [cellText("المجموع", { bold: true }), cellText(stats.totalRegistered, { bold: true }), cellText(stats.totalCompleted, { bold: true }), cellText(stats.globalPct === null ? "—" : stats.globalPct + "%", { bold: true })] }),
    ],
  });

  const teamTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [cellText("المديرية الإقليمية", { bold: true }), cellText("الميسر", { bold: true })] }),
      ...(stats.team.length > 0
        ? stats.team.map((t) => new TableRow({ children: [cellText(provName(t.province)), cellText(t.tutor)] }))
        : [new TableRow({ children: [cellText("—"), cellText("لا يوجد ميسر مسجل")] })]),
    ],
  });

  const doc = new Document({
    sections: [
      {
        children: [
          P(orgLine1),
          P(orgLine2, { spacing: { after: 300 } }),
          P(title, { heading: HeadingLevel.TITLE }),
          P(`الدورة : ${sessionName}`),
          P(`الوحدة : ${moduleName}`, { spacing: { after: 300 } }),
          P(`من إنجاز : ${preparedBy || "—"}`),
          P(`بتنسيق مع : ${coordinatedWith || "—"}`, { spacing: { after: 400 } }),

          P("ملخص عام", { heading: HeadingLevel.HEADING_1 }),
          P(summary, { spacing: { after: 300 } }),

          P("عدد المستفيدين حسب المديرية الإقليمية", { heading: HeadingLevel.HEADING_1, spacing: { before: 200 } }),
          registrationTable,

          P("الفريق الجهوي المكلف بالتأطير", { heading: HeadingLevel.HEADING_1, spacing: { before: 300 } }),
          teamTable,

          P("خلاصة عامة", { heading: HeadingLevel.HEADING_1, spacing: { before: 300 } }),
          P(conclusion),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const safeName = `Rapport_AR_${sessionName}_${moduleName}`.replace(/[^a-zA-Z0-9_\-]+/g, "_");
  saveAs(blob, `${safeName}.docx`);
}

async function generateReportDocx({ orgLine1, orgLine2, title, sessionName, moduleName, preparedBy, coordinatedWith, summary, conclusion, stats }) {
  const cellText = (text, opts = {}) =>
    new TableCell({
      children: [new Paragraph({ children: [new TextRun({ text: String(text), bold: !!opts.bold })] })],
      width: { size: opts.width || 25, type: WidthType.PERCENTAGE },
    });

  const registrationTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [cellText("Direction provinciale", { bold: true }), cellText("Inscrits", { bold: true }), cellText("Ont satisfait le module", { bold: true }), cellText("% de réussite", { bold: true })] }),
      ...stats.rows.map((r) => new TableRow({ children: [cellText(r.province), cellText(r.registered), cellText(r.completed), cellText(r.pct === null ? "—" : r.pct + "%")] })),
      new TableRow({ children: [cellText("Total", { bold: true }), cellText(stats.totalRegistered, { bold: true }), cellText(stats.totalCompleted, { bold: true }), cellText(stats.globalPct === null ? "—" : stats.globalPct + "%", { bold: true })] }),
    ],
  });

  const teamTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [cellText("Direction provinciale", { bold: true }), cellText("Tuteur", { bold: true })] }),
      ...(stats.team.length > 0
        ? stats.team.map((t) => new TableRow({ children: [cellText(t.province), cellText(t.tutor)] }))
        : [new TableRow({ children: [cellText("—"), cellText("Aucun tuteur renseigné")] })]),
    ],
  });

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ text: orgLine1, alignment: AlignmentType.CENTER }),
          new Paragraph({ text: orgLine2, alignment: AlignmentType.CENTER, spacing: { after: 300 } }),
          new Paragraph({ text: title, heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
          new Paragraph({ text: `Session : ${sessionName}`, alignment: AlignmentType.CENTER }),
          new Paragraph({ text: `Module : ${moduleName}`, alignment: AlignmentType.CENTER, spacing: { after: 300 } }),
          new Paragraph({ children: [new TextRun({ text: "Réalisé par : ", bold: true }), new TextRun(preparedBy || "—")] }),
          new Paragraph({ children: [new TextRun({ text: "En coordination avec : ", bold: true }), new TextRun(coordinatedWith || "—")], spacing: { after: 400 } }),

          new Paragraph({ text: "Résumé général", heading: HeadingLevel.HEADING_1 }),
          new Paragraph({ text: summary, spacing: { after: 300 } }),

          new Paragraph({ text: "Bénéficiaires par direction provinciale", heading: HeadingLevel.HEADING_1, spacing: { before: 200 } }),
          registrationTable,

          new Paragraph({ text: "Équipe régionale encadrante", heading: HeadingLevel.HEADING_1, spacing: { before: 300 } }),
          teamTable,

          new Paragraph({ text: "Conclusion générale", heading: HeadingLevel.HEADING_1, spacing: { before: 300 } }),
          new Paragraph({ text: conclusion }),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const safeName = `Rapport_${sessionName}_${moduleName}`.replace(/[^a-zA-Z0-9_\-]+/g, "_");
  saveAs(blob, `${safeName}.docx`);
}

function ReportTab({ sessions, modules, entries }) {
  const [sessionId, setSessionId] = useState("");
  const [moduleId, setModuleId] = useState("");
  const [orgLine1, setOrgLine1] = useState("Direction des Ressources Pédagogiques et Numériques");
  const [orgLine2, setOrgLine2] = useState("Service de l'Apprentissage et de l'Enseignement à Distance");
  const [title, setTitle] = useState("Rapport de suivi de la session de formation e-Qissmi");
  const [preparedBy, setPreparedBy] = useState("");
  const [coordinatedWith, setCoordinatedWith] = useState("");
  const [summary, setSummary] = useState("");
  const [conclusion, setConclusion] = useState("");
  const [generating, setGenerating] = useState(false);

  const [orgLine1Ar, setOrgLine1Ar] = useState("مديرية الموارد البيداغوجية والرقمية");
  const [orgLine2Ar, setOrgLine2Ar] = useState("مصلحة التعلم والتكوين عن بعد");
  const [titleAr, setTitleAr] = useState("تقرير حول سير الدورة التكوينية إ-قسمي");
  const [summaryAr, setSummaryAr] = useState("");
  const [conclusionAr, setConclusionAr] = useState("");
  const [generatingAr, setGeneratingAr] = useState(false);

  const sessionName = sessions.find((s) => s.id === sessionId)?.name || "";
  const moduleName = modules.find((m) => m.id === moduleId)?.name || "";

  const stats = useMemo(() => {
    if (!sessionId || !moduleId) return null;
    return computeReportStats(entries, sessionId, moduleId);
  }, [entries, sessionId, moduleId]);

  const fillDefaults = () => {
    if (!stats) return;
    setSummary(defaultSummary(sessionName, moduleName, stats));
    setConclusion(defaultConclusion(sessionName, moduleName));
  };

  const fillDefaultsAr = () => {
    if (!stats) return;
    setSummaryAr(defaultSummaryAr(sessionName, moduleName, stats));
    setConclusionAr(defaultConclusionAr(sessionName, moduleName));
  };

  const handleGenerate = async () => {
    if (!stats) return;
    setGenerating(true);
    try {
      await generateReportDocx({
        orgLine1, orgLine2, title, sessionName, moduleName, preparedBy, coordinatedWith,
        summary: summary || defaultSummary(sessionName, moduleName, stats),
        conclusion: conclusion || defaultConclusion(sessionName, moduleName),
        stats,
      });
    } finally {
      setGenerating(false);
    }
  };

  const handleGenerateAr = async () => {
    if (!stats) return;
    setGeneratingAr(true);
    try {
      await generateReportDocxArabic({
        orgLine1: orgLine1Ar, orgLine2: orgLine2Ar, title: titleAr, sessionName, moduleName, preparedBy, coordinatedWith,
        summary: summaryAr || defaultSummaryAr(sessionName, moduleName, stats),
        conclusion: conclusionAr || defaultConclusionAr(sessionName, moduleName),
        stats,
      });
    } finally {
      setGeneratingAr(false);
    }
  };

  return (
    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18 }}>
      <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
        <div className="eq-card">
          <h3 style={{ margin: "0 0 4px", fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>Session et module concernés</h3>
          <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.inkSoft }}>Le rapport est généré pour une session et un module précis.</p>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <select className="eq-input" value={sessionId} onChange={(e) => setSessionId(e.target.value)}>
              <option value="">Session…</option>
              {sessions.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            <select className="eq-input" value={moduleId} onChange={(e) => setModuleId(e.target.value)}>
              <option value="">Module…</option>
              {modules.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}
            </select>
          </div>
        </div>

        <div className="eq-card">
          <h3 style={{ margin: "0 0 4px", fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>En-tête du rapport</h3>
          <div style={{ display: "flex", flexDirection: "column", gap: 10, marginTop: 10 }}>
            <input className="eq-input" placeholder="Ligne d'en-tête 1 (direction)" value={orgLine1} onChange={(e) => setOrgLine1(e.target.value)} />
            <input className="eq-input" placeholder="Ligne d'en-tête 2 (service)" value={orgLine2} onChange={(e) => setOrgLine2(e.target.value)} />
            <input className="eq-input" placeholder="Titre du rapport" value={title} onChange={(e) => setTitle(e.target.value)} />
            <input className="eq-input" placeholder="Réalisé par" value={preparedBy} onChange={(e) => setPreparedBy(e.target.value)} />
            <input className="eq-input" placeholder="En coordination avec" value={coordinatedWith} onChange={(e) => setCoordinatedWith(e.target.value)} />
          </div>
        </div>

        <div className="eq-card">
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
            <h3 style={{ margin: 0, fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>Texte du rapport</h3>
            <button className="eq-btn eq-btn-secondary" disabled={!stats} onClick={fillDefaults}>
              <RefreshCw size={13} /> Générer un texte automatique
            </button>
          </div>
          <p style={{ margin: "0 0 10px", fontSize: 12.5, color: C.inkSoft }}>Modifiable librement avant de générer le document.</p>
          <label style={{ fontSize: 12.5, color: C.inkSoft, display: "block", marginBottom: 4 }}>Résumé général</label>
          <textarea className="eq-input" rows={5} value={summary} onChange={(e) => setSummary(e.target.value)} style={{ marginBottom: 10 }} />
          <label style={{ fontSize: 12.5, color: C.inkSoft, display: "block", marginBottom: 4 }}>Conclusion générale</label>
          <textarea className="eq-input" rows={4} value={conclusion} onChange={(e) => setConclusion(e.target.value)} />
        </div>

        <button className="eq-btn eq-btn-primary" disabled={!stats || generating} onClick={handleGenerate} style={{ justifyContent: "center", padding: "12px 14px" }}>
          <FileText size={15} /> {generating ? "Génération…" : "Générer le rapport français (.docx)"}
        </button>

        <div className="eq-card">
          <h3 style={{ margin: "0 0 4px", fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>النسخة العربية</h3>
          <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.inkSoft }}>نفس المعطيات، بصيغة عربية مستقلة</p>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <input className="eq-input" style={{ textAlign: "right", direction: "rtl" }} placeholder="السطر الأول" value={orgLine1Ar} onChange={(e) => setOrgLine1Ar(e.target.value)} />
            <input className="eq-input" style={{ textAlign: "right", direction: "rtl" }} placeholder="السطر الثاني" value={orgLine2Ar} onChange={(e) => setOrgLine2Ar(e.target.value)} />
            <input className="eq-input" style={{ textAlign: "right", direction: "rtl" }} placeholder="عنوان التقرير" value={titleAr} onChange={(e) => setTitleAr(e.target.value)} />
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", margin: "14px 0 4px" }}>
            <label style={{ fontSize: 12.5, color: C.inkSoft }}>ملخص عام</label>
            <button className="eq-btn eq-btn-secondary" disabled={!stats} onClick={fillDefaultsAr}>
              <RefreshCw size={13} /> استخراج نص تلقائي
            </button>
          </div>
          <textarea className="eq-input" style={{ textAlign: "right", direction: "rtl" }} rows={5} value={summaryAr} onChange={(e) => setSummaryAr(e.target.value)} />
          <label style={{ fontSize: 12.5, color: C.inkSoft, display: "block", margin: "10px 0 4px" }}>خلاصة عامة</label>
          <textarea className="eq-input" style={{ textAlign: "right", direction: "rtl" }} rows={4} value={conclusionAr} onChange={(e) => setConclusionAr(e.target.value)} />
        </div>

        <button className="eq-btn eq-btn-primary" disabled={!stats || generatingAr} onClick={handleGenerateAr} style={{ justifyContent: "center", padding: "12px 14px" }}>
          <FileText size={15} /> {generatingAr ? "جاري الإنشاء…" : "استخراج التقرير بالعربية (.docx)"}
        </button>

        {!stats && (
          <div style={{ fontSize: 12, color: C.clay, display: "flex", gap: 6 }}>
            <AlertCircle size={14} style={{ flexShrink: 0, marginTop: 1 }} />
            Choisissez une session et un module pour activer la génération.
          </div>
        )}
      </div>

      <div className="eq-card">
        <h3 style={{ margin: "0 0 10px", fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 16 }}>Aperçu des données calculées</h3>
        {!stats ? (
          <div style={{ color: C.inkSoft, fontSize: 13 }}>Sélectionnez une session et un module pour voir l'aperçu.</div>
        ) : (
          <>
            <table className="eq-table" style={{ marginBottom: 16 }}>
              <thead>
                <tr><th>Direction</th><th>Inscrits</th><th>Ont satisfait</th><th>%</th></tr>
              </thead>
              <tbody>
                {stats.rows.map((r) => (
                  <tr key={r.province}>
                    <td>{r.province}</td>
                    <td>{r.registered}</td>
                    <td>{r.completed}</td>
                    <td>{r.pct === null ? "—" : `${r.pct}%`}</td>
                  </tr>
                ))}
                <tr style={{ fontWeight: 700 }}>
                  <td>Total</td>
                  <td>{stats.totalRegistered}</td>
                  <td>{stats.totalCompleted}</td>
                  <td>{stats.globalPct === null ? "—" : `${stats.globalPct}%`}</td>
                </tr>
              </tbody>
            </table>
            <div style={{ fontSize: 12.5, color: C.inkSoft, marginBottom: 6 }}>{stats.team.length} tuteur(s) identifié(s) pour cette session</div>
            <table className="eq-table">
              <thead>
                <tr><th>Direction</th><th>Tuteur</th></tr>
              </thead>
              <tbody>
                {stats.team.map((t) => (
                  <tr key={t.province + t.tutor}>
                    <td>{t.province}</td>
                    <td>{t.tutor}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </>
        )}
      </div>
    </div>
  );
}
