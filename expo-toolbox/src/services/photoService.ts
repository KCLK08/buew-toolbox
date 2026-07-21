import { photoRepository } from '../repositories';
import { persistLocalFile } from '../storage/fileService';
import { nowIso } from '../lib/ids';
import type { ParentType, Photo } from '../types/offline';

export async function savePhotoOffline(input: {
  sourceUri: string;
  mimeType?: string;
  parentId?: string | null;
  parentType?: ParentType | null;
  /** @deprecated Use parentId/parentType */
  projectId?: string | null;
  /** @deprecated Use parentId/parentType */
  diaryRunId?: string | null;
  /** @deprecated Use parentId/parentType */
  defectId?: string | null;
}): Promise<Photo> {
  const stored = await persistLocalFile({
    sourceUri: input.sourceUri,
    mimeType: input.mimeType ?? 'image/jpeg',
    kind: 'photo'
  });

  let parentId = input.parentId ?? null;
  let parentType = input.parentType ?? null;
  if (!parentId || !parentType) {
    if (input.defectId) {
      parentId = input.defectId;
      parentType = 'defect';
    } else if (input.diaryRunId) {
      parentId = input.diaryRunId;
      parentType = 'diary_entry';
    } else if (input.projectId) {
      parentId = input.projectId;
      parentType = 'project';
    }
  }

  const filename = stored.relativePath.includes('/')
    ? stored.relativePath.slice(stored.relativePath.lastIndexOf('/') + 1)
    : stored.relativePath;

  return photoRepository.addPhoto({
    id: stored.id,
    parent_id: parentId,
    parent_type: parentType,
    filename,
    local_path: stored.relativePath,
    mime_type: stored.mimeType,
    file_size: stored.byteSize,
    status: 'ready',
    created_at: stored.createdAt || nowIso()
  });
}
