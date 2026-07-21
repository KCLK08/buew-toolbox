import { Link } from 'react-router-dom';
import { AUTH_COPY, useAuth } from '@buew/shared';

import { TOOLS } from '../lib/tools';
import { webConfig } from '../lib/supabase';

export function DashboardPage() {
  const { user, signOut } = useAuth();

  const onLogout = async () => {
    await signOut();
  };

  return (
    <div className="page">
      <header className="topbar">
        <div>
          <p className="brand">BÜW-Toolbox</p>
          <p className="muted">
            {user?.displayName ?? user?.email} · Rolle: {user?.role ?? 'benutzer'}
          </p>
        </div>
        <div className="topbar-actions">
          <Link className="ghost-btn" to="/settings">
            Einstellungen
          </Link>
          <button className="ghost-btn" type="button" onClick={() => void onLogout()}>
            {AUTH_COPY.logout}
          </button>
        </div>
      </header>

      <section className="hero">
        <h1>BÜW-Toolbox</h1>
        <p className="sub">Zentrale Übersicht für digitale Baustellen‑Workflows und Dokumentation.</p>
      </section>

      <section className="grid">
        {TOOLS.map((tool) => (
          <a
            key={tool.id}
            className="card-link"
            href={`${webConfig.toolboxWebBaseUrl.replace(/\/$/, '')}${tool.href}`}
          >
            <article className="card">
              <div className="card-header">
                <div className="icon">
                  <img src={tool.icon} alt={`${tool.title} Icon`} />
                </div>
                <h2>{tool.title}</h2>
              </div>
              <p>{tool.description}</p>
            </article>
          </a>
        ))}
      </section>
    </div>
  );
}
