import * as FileSystem from 'expo-file-system/legacy';
import * as ImageManipulator from 'expo-image-manipulator';
import * as ImagePicker from 'expo-image-picker';

const PHOTO_ROOT = `${FileSystem.documentDirectory}sitereport/photos/`;

async function ensurePhotoDir(protocolId: string): Promise<string> {
  const dir = `${PHOTO_ROOT}${protocolId}/`;
  await FileSystem.makeDirectoryAsync(dir, { intermediates: true });
  return dir;
}

export async function compressAndPersistPhoto(
  sourceUri: string,
  protocolId: string,
  entryId: string
): Promise<string> {
  const manipulated = await ImageManipulator.manipulateAsync(
    sourceUri,
    [{ resize: { width: 1600 } }],
    { compress: 0.75, format: ImageManipulator.SaveFormat.JPEG }
  );
  const dir = await ensurePhotoDir(protocolId);
  const localPath = `${dir}${entryId}.jpg`;
  await FileSystem.copyAsync({ from: manipulated.uri, to: localPath });
  return localPath;
}

export async function captureProtocolPhoto(protocolId: string, entryId: string): Promise<string | null> {
  const permission = await ImagePicker.requestCameraPermissionsAsync();
  if (!permission.granted) {
    throw new Error('Kamerazugriff ist erforderlich.');
  }
  const result = await ImagePicker.launchCameraAsync({ quality: 0.8 });
  if (result.canceled || !result.assets[0]?.uri) {
    return null;
  }
  return compressAndPersistPhoto(result.assets[0].uri, protocolId, entryId);
}

export async function deleteEntryPhoto(photoPath: string | null): Promise<void> {
  if (!photoPath) return;
  await FileSystem.deleteAsync(photoPath, { idempotent: true });
}

export async function deleteProtocolPhotos(protocolId: string): Promise<void> {
  const dir = `${PHOTO_ROOT}${protocolId}/`;
  await FileSystem.deleteAsync(dir, { idempotent: true });
}
