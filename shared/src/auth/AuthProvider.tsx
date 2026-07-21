import {
  createContext,
  createElement,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
  type PropsWithChildren,
  type ReactElement
} from 'react';
import type { Session } from '@supabase/supabase-js';

import { AUTH_COPY } from '../constants/theme';
import type { NetworkMonitor } from '../lib/network';
import { isOfflineError } from '../lib/errors';
import type { AuthService } from '../services/authService';
import type {
  AuthResult,
  AuthUserView,
  ResetPasswordInput,
  SignInInput,
  SignUpInput
} from '../types/auth';

export type AuthContextValue = {
  ready: boolean;
  online: boolean;
  session: Session | null;
  user: AuthUserView | null;
  isAuthenticated: boolean;
  signIn: (input: SignInInput) => Promise<AuthResult>;
  signUp: (input: SignUpInput) => Promise<AuthResult>;
  resetPassword: (input: ResetPasswordInput) => Promise<AuthResult>;
  signOut: () => Promise<AuthResult>;
  refreshUser: () => Promise<void>;
};

const AuthContext = createContext<AuthContextValue | null>(null);

export type AuthProviderProps = PropsWithChildren<{
  authService: AuthService;
  networkMonitor: NetworkMonitor;
}>;

export function AuthProvider({
  authService,
  networkMonitor,
  children
}: AuthProviderProps): ReactElement {
  const [ready, setReady] = useState(false);
  const [online, setOnline] = useState(true);
  const [session, setSession] = useState<Session | null>(null);
  const [user, setUser] = useState<AuthUserView | null>(null);

  useEffect(() => {
    let active = true;

    networkMonitor.getStatus().then((status) => {
      if (active) setOnline(status.online);
    });

    const unsubscribeNetwork = networkMonitor.subscribe((status) => {
      setOnline(status.online);
    });

    authService
      .getSession()
      .then(async (current) => {
        if (!active) return;
        setSession(current);
        setUser(await authService.resolveUser(current));
        setReady(true);
      })
      .catch(() => {
        if (!active) return;
        setSession(null);
        setUser(null);
        setReady(true);
      });

    const unsubscribeAuth = authService.onAuthStateChange((nextSession) => {
      setSession(nextSession);
      void authService.resolveUser(nextSession).then((resolved) => {
        if (active) setUser(resolved);
      });
    });

    return () => {
      active = false;
      unsubscribeNetwork();
      unsubscribeAuth();
    };
  }, [authService, networkMonitor]);

  const guardOffline = useCallback((): AuthResult | null => {
    if (!online) {
      return { error: AUTH_COPY.offline };
    }
    return null;
  }, [online]);

  const wrapNetwork = useCallback(
    async (action: () => Promise<AuthResult>): Promise<AuthResult> => {
      const offline = guardOffline();
      if (offline) return offline;
      try {
        return await action();
      } catch (error) {
        if (isOfflineError(error)) {
          return { error: AUTH_COPY.offline };
        }
        return { error: error instanceof Error ? error.message : 'Unbekannter Fehler' };
      }
    },
    [guardOffline]
  );

  const signIn = useCallback(
    (input: SignInInput) => wrapNetwork(() => authService.signIn(input)),
    [authService, wrapNetwork]
  );

  const signUp = useCallback(
    (input: SignUpInput) => wrapNetwork(() => authService.signUp(input)),
    [authService, wrapNetwork]
  );

  const resetPassword = useCallback(
    (input: ResetPasswordInput) => wrapNetwork(() => authService.resetPassword(input)),
    [authService, wrapNetwork]
  );

  const signOut = useCallback(
    () => wrapNetwork(() => authService.signOut()),
    [authService, wrapNetwork]
  );

  const refreshUser = useCallback(async () => {
    const current = await authService.getSession();
    setSession(current);
    setUser(await authService.resolveUser(current));
  }, [authService]);

  const value = useMemo<AuthContextValue>(
    () => ({
      ready,
      online,
      session,
      user,
      isAuthenticated: Boolean(session?.user),
      signIn,
      signUp,
      resetPassword,
      signOut,
      refreshUser
    }),
    [ready, online, session, user, signIn, signUp, resetPassword, signOut, refreshUser]
  );

  return createElement(AuthContext.Provider, { value }, children);
}

export function useAuth(): AuthContextValue {
  const ctx = useContext(AuthContext);
  if (!ctx) {
    throw new Error('useAuth must be used within AuthProvider');
  }
  return ctx;
}
