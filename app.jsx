import React, { useState } from "react";
import { 
  Users, FileText, Plus, Trash2, CalendarDays, UserPlus 
} from "lucide-react";

const CSS = `
  :root { --bg: #f8fafc; --panel: #ffffff; --primary: #4f46e5; --text: #1e293b; }
  body { background: var(--bg); color: var(--text); font-family: 'Inter', system-ui, sans-serif; margin: 0; padding: 20px; }
  .app-shell { max-width: 1200px; margin: 0 auto; display: flex; flex-direction: column; gap: 20px; }
  .header { display: flex; justify-content: space-between; align-items: center; background: white; padding: 20px; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
  .tabs { display: flex; gap: 10px; }
  .tab-btn { padding: 10px 20px; border: none; border-radius: 8px; cursor: pointer; display: flex; align-items: center; gap: 8px; font-weight: 500; transition: 0.2s; }
  .tab-btn-active { background: var(--primary); color: white; }
  .panel { background: white; padding: 24px; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); min-height: 500px; }
  .card { border: 1px solid #e2e8f0; padding: 16px; border-radius: 8px; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center; }
  .input-field { padding: 8px; border: 1px solid #cbd5e1; border-radius: 6px; margin-right: 10px; }
  .add-btn { background: var(--primary); color: white; padding: 8px 16px; border-radius: 6px; border: none; cursor: pointer; }
`;

export default function App() {
  const [tab, setTab] = useState("dashboard");
  const [collaborateurs, setCollaborateurs] = useState([
    { id: 1, nom: "Jean Dupont" },
    { id: 2, nom: "Marie Curie" }
  ]);
  const [nouveauNom, setNouveauNom] = useState("");

  const ajouterCollaborateur = () => {
    if (!nouveauNom.trim()) return;
    setCollaborateurs([...collaborateurs, { id: Date.now(), nom: nouveauNom }]);
    setNouveauNom("");
  };

  const supprimerCollaborateur = (id) => {
    setCollaborateurs(collaborateurs.filter(c => c.id !== id));
  };

  return (
    <div className="app-shell">
      <style>{CSS}</style>
      <header className="header">
        <div style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
          {}
          <img 
            src="https://placehold.co/40x40/4f46e5/white?text=RD" 
            alt="Logo" 
            style={{ width: '40px', height: '40px', borderRadius: '8px' }} 
          />
          <h1 style={{ margin: 0, fontSize: '1.5rem' }}>Registre Disciplinaire</h1>
        </div>
        <nav className="tabs">
          <button className={`tab-btn ${tab === 'dashboard' ? 'tab-btn-active' : ''}`} onClick={() => setTab('dashboard')}>
            <CalendarDays size={16} /> Aujourd'hui
          </button>
          <button className={`tab-btn ${tab === 'collaborateurs' ? 'tab-btn-active' : ''}`} onClick={() => setTab('collaborateurs')}>
            <Users size={16} /> Collaborateurs
          </button>
        </nav>
      </header>

      <main className="panel">
        {tab === 'dashboard' && (
          <div>
            <h2>Suivi en temps réel</h2>
            <p>Aucune procédure disciplinaire en cours pour aujourd'hui.</p>
          </div>
        )}
        {tab === 'collaborateurs' && (
          <div>
            <h2>Gestion des collaborateurs</h2>
            <div style={{ marginBottom: '20px' }}>
              <input 
                className="input-field"
                value={nouveauNom} 
                onChange={(e) => setNouveauNom(e.target.value)}
                placeholder="Nom du collaborateur..."
              />
              <button className="add-btn" onClick={ajouterCollaborateur}>
                <Plus size={16} style={{ display: 'inline', marginRight: '5px' }} /> Ajouter
              </button>
            </div>
            {collaborateurs.map(c => (
              <div key={c.id} className="card">
                <span>{c.nom}</span>
                <button onClick={() => supprimerCollaborateur(c.id)} style={{ border: 'none', background: 'none', color: '#ef4444', cursor: 'pointer' }}>
                  <Trash2 size={18} />
                </button>
              </div>
            ))}
          </div>
        )}
      </main>
    </div>
  );
}
