import * as Crypto from 'expo-crypto';

export function nowIso(): string {
  return new Date().toISOString();
}

export async function createUuid(): Promise<string> {
  return Crypto.randomUUID();
}

export function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0;
}
