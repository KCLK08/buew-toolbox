import { Link } from 'react-router-dom';
import { useAuth } from '@buew/shared';

export function SettingsPage() {
  const { user } = useAuth();

  return (
    <div className="page narrow">
      <header className="topbar">
        <Link className="ghost-btn" to="/">
          ← Toolbox
        </Link>
      </header>
      <section className="panel">
        <h1>Einstellungen</h1>
        <p className="muted">Profil und Konto</p>
        <dl className="details">
          <div>
            <dt>Name</dt>
            <dd>{user?.displayName ?? '—'}</dd>
          </div>
          <div>
            <dt>E-Mail</dt>
            <dd>{user?.email ?? '—'}</dd>
          </div>
          <div>
            <dt>Rolle</dt>
            <dd>{user?.role ?? 'benutzer'}</dd>
          </div>
        </dl>
        <p className="hint">
          Biometrische Anmeldung ist nur in der Expo-App verfügbar.
        </p>
      </section>
    </div>
  );
}
