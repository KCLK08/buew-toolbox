export type UserRole = 'admin' | 'benutzer';

export const USER_ROLES = ['admin', 'benutzer'] as const;

export const DEFAULT_USER_ROLE: UserRole = 'benutzer';

export type Profile = {
  id: string;
  email: string | null;
  first_name: string | null;
  last_name: string | null;
  display_name: string | null;
  role: UserRole;
  created_at: string;
  updated_at: string;
};

export type AuthUserView = {
  id: string;
  email: string | null;
  firstName: string | null;
  lastName: string | null;
  displayName: string | null;
  role: UserRole;
};

export type AuthResult = {
  error: string | null;
};

export type SignUpInput = {
  firstName: string;
  lastName: string;
  email: string;
  password: string;
};

export type SignInInput = {
  email: string;
  password: string;
};

export type ResetPasswordInput = {
  email: string;
};
