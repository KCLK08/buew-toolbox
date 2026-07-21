import * as SecureStore from 'expo-secure-store';

const CHUNK_SIZE = 1800;

async function setItem(key: string, value: string): Promise<void> {
  if (value.length <= CHUNK_SIZE) {
    await SecureStore.setItemAsync(key, value);
    await SecureStore.deleteItemAsync(`${key}_chunks`).catch(() => undefined);
    return;
  }

  const chunkCount = Math.ceil(value.length / CHUNK_SIZE);
  await SecureStore.setItemAsync(`${key}_chunks`, String(chunkCount));
  const writes: Promise<void>[] = [];
  for (let i = 0; i < chunkCount; i += 1) {
    const chunk = value.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
    writes.push(SecureStore.setItemAsync(`${key}_${i}`, chunk));
  }
  await Promise.all(writes);
  await SecureStore.deleteItemAsync(key).catch(() => undefined);
}

async function getItem(key: string): Promise<string | null> {
  const chunkCountRaw = await SecureStore.getItemAsync(`${key}_chunks`);
  if (!chunkCountRaw) {
    return SecureStore.getItemAsync(key);
  }

  const chunkCount = Number(chunkCountRaw);
  if (!Number.isFinite(chunkCount) || chunkCount <= 0) {
    return null;
  }

  const chunks: string[] = [];
  for (let i = 0; i < chunkCount; i += 1) {
    const chunk = await SecureStore.getItemAsync(`${key}_${i}`);
    if (chunk == null) return null;
    chunks.push(chunk);
  }
  return chunks.join('');
}

async function removeItem(key: string): Promise<void> {
  const chunkCountRaw = await SecureStore.getItemAsync(`${key}_chunks`);
  if (chunkCountRaw) {
    const chunkCount = Number(chunkCountRaw);
    const deletes: Promise<void>[] = [SecureStore.deleteItemAsync(`${key}_chunks`)];
    for (let i = 0; i < chunkCount; i += 1) {
      deletes.push(SecureStore.deleteItemAsync(`${key}_${i}`));
    }
    await Promise.all(deletes.map((p) => p.catch(() => undefined)));
  }
  await SecureStore.deleteItemAsync(key).catch(() => undefined);
}

/** SecureStore adapter for Supabase Auth (chunked for large sessions). */
export const secureAuthStorage = {
  getItem,
  setItem,
  removeItem
};
