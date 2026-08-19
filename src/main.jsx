import React, { useEffect, useState } from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import "./index.css";
import { auth } from "./firebase.js";
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from "firebase/auth";

const C = { teal: "#0E5C55", tealDark: "#0A423D", bg: "#F6F1E7", ink: "#20302C", inkSoft: "#5B6B63", line: "#E4DBC7", bad: "#B24B32" };

function Root() {
  const [user, setUser] = useState(undefined); // undefined = chargement, null = déconnecté
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (u) => setUser(u));
    return unsubscribe;
  }, []);

  const handleLogin = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);
    try {
      await signInWithEmailAndPassword(auth, email, password);
    } catch (err) {
      setError("Email ou mot de passe incorrect.");
    } finally {
      setLoading(false);
    }
  };

  if (user === undefined) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", fontFamily: "ui-sans-serif, system-ui", color: C.inkSoft }}>
        Chargement…
      </div>
    );
  }

  if (!user) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", background: C.bg, fontFamily: "ui-sans-serif, system-ui, sans-serif" }}>
        <form onSubmit={handleLogin} style={{ background: "#fff", border: `1px solid ${C.line}`, borderRadius: 12, padding: 32, width: 320, display: "flex", flexDirection: "column", gap: 12 }}>
          <div style={{ fontSize: 11.5, fontWeight: 700, letterSpacing: "0.08em", color: C.teal, textTransform: "uppercase" }}>e-Qissmi</div>
          <h2 style={{ margin: "0 0 8px", fontFamily: "ui-serif, Georgia, serif", color: C.tealDark, fontSize: 20 }}>Connexion</h2>
          <input
            type="email" placeholder="Email" value={email} required
            onChange={(e) => setEmail(e.target.value)}
            style={{ border: `1px solid ${C.line}`, borderRadius: 8, padding: "9px 11px", fontSize: 14 }}
          />
          <input
            type="password" placeholder="Mot de passe" value={password} required
            onChange={(e) => setPassword(e.target.value)}
            style={{ border: `1px solid ${C.line}`, borderRadius: 8, padding: "9px 11px", fontSize: 14 }}
          />
          {error && <div style={{ color: C.bad, fontSize: 13 }}>{error}</div>}
          <button
            type="submit" disabled={loading}
            style={{ background: C.teal, color: "#fff", border: "none", borderRadius: 8, padding: "10px 14px", fontSize: 14, fontWeight: 600, cursor: "pointer" }}
          >
            {loading ? "Connexion…" : "Se connecter"}
          </button>
        </form>
      </div>
    );
  }

  return <App user={user} onLogout={() => signOut(auth)} />;
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <Root />
  </React.StrictMode>
);
