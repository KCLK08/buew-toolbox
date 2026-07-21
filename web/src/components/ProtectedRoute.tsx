import { Navigate, Outlet, useLocation } from 'react-router-dom';
import { useAuth } from '@buew/shared';

export function ProtectedRoute() {
  const { ready, isAuthenticated } = useAuth();
  const location = useLocation();

  if (!ready) {
    return <div className="center-screen">Lade Sitzung…</div>;
  }

  if (!isAuthenticated) {
    return <Navigate to="/login" replace state={{ from: location.pathname }} />;
  }

  return <Outlet />;
}

export function GuestRoute() {
  const { ready, isAuthenticated } = useAuth();

  if (!ready) {
    return <div className="center-screen">Lade Sitzung…</div>;
  }

  if (isAuthenticated) {
    return <Navigate to="/" replace />;
  }

  return <Outlet />;
}
