import * as FileSystem from 'expo-file-system/legacy';
import * as ImageManipulator from 'expo-image-manipulator';
import * as ImagePicker from 'expo-image-picker';

import { savePhotoAsset } from '../db/database';
import type { PhotoDocMeta } from '../types';
import { nowIso } from '../../../lib/ids';

export async function capturePhotoDocEntry(runId: string): Promise<PhotoDocMeta['entries'][number] | null> {
  const permission = await ImagePicker.requestCameraPermissionsAsync();
  if (!permission.granted) {
    throw new Error('Kamerazugriff ist erforderlich.');
  }

  const result = await ImagePicker.launchCameraAsync({ quality: 0.8 });
  if (result.canceled || !result.assets[0]?.uri) {
    return null;
  }

  const manipulated = await ImageManipulator.manipulateAsync(
    result.assets[0].uri,
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

  return {
    id: entryId,
    createdAt: nowIso(),
    mimeType: 'image/jpeg',
    localPath
  };
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
