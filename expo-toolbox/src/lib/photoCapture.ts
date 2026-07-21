import * as ImagePicker from 'expo-image-picker';
import { Alert } from 'react-native';

import { savePhotoOffline } from '../services/photoService';
import type { ParentType, Photo } from '../types/offline';

async function ensureCameraPermission(): Promise<boolean> {
  const current = await ImagePicker.getCameraPermissionsAsync();
  if (current.granted) return true;
  const requested = await ImagePicker.requestCameraPermissionsAsync();
  return requested.granted;
}

export async function captureAndSavePhoto(input: {
  parentId: string;
  parentType: ParentType;
}): Promise<Photo | null> {
  const granted = await ensureCameraPermission();
  if (!granted) {
    Alert.alert('Kamera', 'Kamerazugriff ist erforderlich, um Fotos aufzunehmen.');
    return null;
  }

  const result = await ImagePicker.launchCameraAsync({
    mediaTypes: ['images'],
    quality: 0.8,
    exif: false
  });

  if (result.canceled || !result.assets?.[0]?.uri) {
    return null;
  }

  const asset = result.assets[0];
  return savePhotoOffline({
    sourceUri: asset.uri,
    mimeType: asset.mimeType ?? 'image/jpeg',
    parentId: input.parentId,
    parentType: input.parentType
  });
}
