import { useAuth } from '../auth/AuthProvider';

export function useRequireAuth(): ReturnType<typeof useAuth> {
  const auth = useAuth();
  return auth;
}

export function useIsAdmin(): boolean {
  const { user } = useAuth();
  return user?.role === 'admin';
}
