import * as FileSystem from 'expo-file-system/legacy';
import * as ImageManipulator from 'expo-image-manipulator';
import * as ImagePicker from 'expo-image-picker';

import { deletePhotoAsset, savePhotoAsset } from '../db/database';
import type { PhotoDocMeta } from '../types';
import { nowIso } from '../../../lib/ids';
import { requestDatabaseBackup } from '../../../storage/backupService';

type PhotoSource = 'camera' | 'gallery';

async function persistPhotoFromUri(
  runId: string,
  sourceUri: string
): Promise<PhotoDocMeta['entries'][number]> {
  const manipulated = await ImageManipulator.manipulateAsync(
    sourceUri,
    [{ resize: { width: 1600 } }],
    { compress: 0.75, format: ImageManipulator.SaveFormat.JPEG }
  );

  const entryId = `photo_${Date.now()}`;
  const photoDir = `${FileSystem.documentDirectory}bautagebuch/photos/${runId}/`;
  await FileSystem.makeDirectoryAsync(photoDir, { intermediates: true });
  const localPath = `${photoDir}${entryId}.jpg`;
  await FileSystem.copyAsync({ from: manipulated.uri, to: localPath });

  const info = await FileSystem.getInfoAsync(localPath);
  const sizeBytes = info.exists && 'size' in info ? Number(info.size || 0) : 0;
  await savePhotoAsset({
    runId,
    entryId,
    localPath,
    mimeType: 'image/jpeg',
    sizeBytes
  });

  void requestDatabaseBackup('photo_added');

  return {
    id: entryId,
    createdAt: nowIso(),
    mimeType: 'image/jpeg',
    localPath
  };
}

async function pickImage(source: PhotoSource): Promise<string | null> {
  if (source === 'camera') {
    const permission = await ImagePicker.requestCameraPermissionsAsync();
    if (!permission.granted) {
      throw new Error('Kamerazugriff ist erforderlich.');
    }
    const result = await ImagePicker.launchCameraAsync({ quality: 0.8 });
    if (result.canceled || !result.assets[0]?.uri) return null;
    return result.assets[0].uri;
  }

  const permission = await ImagePicker.requestMediaLibraryPermissionsAsync();
  if (!permission.granted) {
    throw new Error('Zugriff auf die Fotobibliothek ist erforderlich.');
  }
  const result = await ImagePicker.launchImageLibraryAsync({
    quality: 0.8,
    mediaTypes: ['images']
  });
  if (result.canceled || !result.assets[0]?.uri) return null;
  return result.assets[0].uri;
}

export async function capturePhotoDocEntry(runId: string): Promise<PhotoDocMeta['entries'][number] | null> {
  const uri = await pickImage('camera');
  if (!uri) return null;
  return persistPhotoFromUri(runId, uri);
}

export async function pickPhotoDocEntry(runId: string): Promise<PhotoDocMeta['entries'][number] | null> {
  const uri = await pickImage('gallery');
  if (!uri) return null;
  return persistPhotoFromUri(runId, uri);
}

export async function removePhotoDocEntry(
  runId: string,
  entryId: string,
  localPath?: string
): Promise<void> {
  await deletePhotoAsset(runId, entryId, localPath);
  void requestDatabaseBackup('record_deleted');
}

export async function readPhotoBytes(localPath: string): Promise<Uint8Array> {
  const base64 = await FileSystem.readAsStringAsync(localPath, {
    encoding: FileSystem.EncodingType.Base64
  });
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}
