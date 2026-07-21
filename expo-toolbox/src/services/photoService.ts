import { photoRepository } from '../repositories';
import { persistLocalFile } from '../storage/fileService';
import { nowIso } from '../lib/ids';
import type { PhotoRecord } from '../types/offline';

export async function savePhotoOffline(input: {
  sourceUri: string;
  mimeType?: string;
  projectId?: string | null;
  diaryRunId?: string | null;
  defectId?: string | null;
}): Promise<PhotoRecord> {
  const stored = await persistLocalFile({
    sourceUri: input.sourceUri,
    mimeType: input.mimeType ?? 'image/jpeg',
    kind: 'photo'
  });

  return photoRepository.saveMetadata({
    id: stored.id,
    project_id: input.projectId ?? null,
    diary_run_id: input.diaryRunId ?? null,
    defect_id: input.defectId ?? null,
    file_path: stored.relativePath,
    mime_type: stored.mimeType,
    byte_size: stored.byteSize,
    status: 'ready',
    created_at: stored.createdAt || nowIso()
  });
}
