import type { Session, SupabaseClient, User } from '@supabase/supabase-js';

import { DEFAULT_USER_ROLE } from '../types/auth';
import type {
  AuthResult,
  AuthUserView,
  Profile,
  ResetPasswordInput,
  SignInInput,
  SignUpInput,
  UserRole
} from '../types/auth';
import type { Database } from '../types/database';
import { mapAuthError } from '../lib/errors';

export type AuthClient = SupabaseClient<Database>;

export type AuthServiceOptions = {
  client: AuthClient;
  emailRedirectTo?: string;
  passwordResetRedirectTo?: string;
};

function toRole(value: string | null | undefined): UserRole {
  return value === 'admin' ? 'admin' : DEFAULT_USER_ROLE;
}

function profileToView(profile: Profile | null, user: User | null): AuthUserView | null {
  if (!user) return null;
  const meta = user.user_metadata ?? {};

  return {
    id: user.id,
    email: user.email ?? profile?.email ?? null,
    firstName:
      profile?.first_name ??
      (typeof meta.first_name === 'string' ? meta.first_name : null),
    lastName:
      profile?.last_name ??
      (typeof meta.last_name === 'string' ? meta.last_name : null),
    displayName:
      profile?.display_name ??
      (typeof meta.display_name === 'string' ? meta.display_name : null),
    role: profile?.role ?? toRole(typeof meta.role === 'string' ? meta.role : null)
  };
}

export function createAuthService({
  client,
  emailRedirectTo,
  passwordResetRedirectTo
}: AuthServiceOptions) {
  async function fetchProfile(userId: string): Promise<Profile | null> {
    const { data, error } = await client
      .from('profiles')
      .select('*')
      .eq('id', userId)
      .maybeSingle();

    if (error) {
      throw error;
    }

    return data;
  }

  async function resolveUser(session: Session | null): Promise<AuthUserView | null> {
    if (!session?.user) return null;
    try {
      const profile = await fetchProfile(session.user.id);
      return profileToView(profile, session.user);
    } catch {
      return profileToView(null, session.user);
    }
  }

  return {
    async getSession(): Promise<Session | null> {
      const { data, error } = await client.auth.getSession();
      if (error) throw error;
      return data.session;
    },

    onAuthStateChange(callback: (session: Session | null) => void) {
      const { data } = client.auth.onAuthStateChange((_event, session) => {
        callback(session);
      });
      return () => data.subscription.unsubscribe();
    },

    async getCurrentUser(): Promise<AuthUserView | null> {
      const session = await this.getSession();
      return resolveUser(session);
    },

    resolveUser,

    async signIn({ email, password }: SignInInput): Promise<AuthResult> {
      const { error } = await client.auth.signInWithPassword({
        email: email.trim(),
        password
      });
      return { error: error ? mapAuthError(error) : null };
    },

    async signUp({ firstName, lastName, email, password }: SignUpInput): Promise<AuthResult> {
      const displayName = `${firstName.trim()} ${lastName.trim()}`.trim();
      const { error } = await client.auth.signUp({
        email: email.trim(),
        password,
        options: {
          emailRedirectTo,
          data: {
            first_name: firstName.trim(),
            last_name: lastName.trim(),
            display_name: displayName,
            role: DEFAULT_USER_ROLE
          }
        }
      });
      return { error: error ? mapAuthError(error) : null };
    },

    async resetPassword({ email }: ResetPasswordInput): Promise<AuthResult> {
      const { error } = await client.auth.resetPasswordForEmail(email.trim(), {
        redirectTo: passwordResetRedirectTo
      });
      return { error: error ? mapAuthError(error) : null };
    },

    async signOut(): Promise<AuthResult> {
      const { error } = await client.auth.signOut();
      return { error: error ? mapAuthError(error) : null };
    }
  };
}

export type AuthService = ReturnType<typeof createAuthService>;
