import type { AuthError } from '@supabase/supabase-js';

const AUTH_ERROR_MAP: Record<string, string> = {
  invalid_credentials: 'E-Mail oder Passwort ist falsch.',
  email_not_confirmed: 'Bitte bestätige zuerst deine E-Mail-Adresse.',
  user_already_exists: 'Für diese E-Mail existiert bereits ein Konto.',
  over_request_rate_limit: 'Zu viele Versuche. Bitte kurz warten.',
  weak_password: 'Das Passwort erfüllt nicht die Sicherheitsanforderungen.',
  validation_failed: 'Eingaben sind ungültig.',
  network_error: 'Netzwerkfehler. Bitte Verbindung prüfen.'
};

export function mapAuthError(error: AuthError | Error | null | undefined): string {
  if (!error) {
    return 'Unbekannter Fehler';
  }

  const code = 'code' in error && typeof error.code === 'string' ? error.code : '';
  if (code && AUTH_ERROR_MAP[code]) {
    return AUTH_ERROR_MAP[code];
  }

  const message = error.message.toLowerCase();

  if (message.includes('invalid login credentials')) {
    return AUTH_ERROR_MAP.invalid_credentials;
  }
  if (message.includes('email not confirmed')) {
    return AUTH_ERROR_MAP.email_not_confirmed;
  }
  if (message.includes('user already registered') || message.includes('already been registered')) {
    return AUTH_ERROR_MAP.user_already_exists;
  }
  if (message.includes('network') || message.includes('fetch')) {
    return AUTH_ERROR_MAP.network_error;
  }

  return error.message || 'Unbekannter Fehler';
}

export function isOfflineError(error: unknown): boolean {
  if (!error || typeof error !== 'object') return false;
  const message = 'message' in error && typeof error.message === 'string' ? error.message.toLowerCase() : '';
  return message.includes('network') || message.includes('failed to fetch') || message.includes('offline');
}
