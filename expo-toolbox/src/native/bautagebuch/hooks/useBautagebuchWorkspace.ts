import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { useFocusEffect } from 'expo-router';

import {
  createRun,
  deleteRunCascade,
  listRuns,
  renameRun,
  updateRun
} from '../db/database';
import { applyRunDefaultsFromModel } from '../lib/run-defaults';
import { buildBtbTitle } from '../lib/btb-naming';
import { groupRunsByCalendar } from '../lib/group-runs-by-calendar';
import { getActiveTemplateBundle } from '../services/templateService';
import type { BautagebuchRun } from '../types';

export function useBautagebuchWorkspace() {
  const [initialLoading, setInitialLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [runs, setRuns] = useState<BautagebuchRun[]>([]);
  const [templateId, setTemplateId] = useState('');
  const [templateName, setTemplateName] = useState('');
  const [templateReady, setTemplateReady] = useState(false);
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const isFirstFocus = useRef(true);

  const load = useCallback(async (mode: 'initial' | 'refresh' = 'initial') => {
    if (mode === 'initial') setInitialLoading(true);
    else setRefreshing(true);
    setError(null);
    try {
      const bundle = await getActiveTemplateBundle();
      setTemplateId(bundle.template.templateId);
      setTemplateName(bundle.template.templateName);
      setTemplateReady(bundle.template.status === 'ready');
      setSetupModel(bundle.setupModel);
      setRuns(await listRuns());
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Bautagebuch konnte nicht geladen werden.');
    } finally {
      if (mode === 'initial') setInitialLoading(false);
      else setRefreshing(false);
    }
  }, []);

  useEffect(() => {
    void load('initial');
  }, [load]);

  useFocusEffect(
    useCallback(() => {
      if (isFirstFocus.current) {
        isFirstFocus.current = false;
        return;
      }
      void load('refresh');
    }, [load])
  );

  const runTree = useMemo(
    () => groupRunsByCalendar(runs, { setupModel }),
    [runs, setupModel]
  );

  const createNewRun = useCallback(
    async (newName: string) => {
      if (!templateId || !setupModel || !templateReady) {
        throw new Error('Vorlage nicht bereit');
      }
      const title = buildBtbTitle(newName);
      let run = await createRun({
        templateId,
        title,
        setupVersion: Number(setupModel.version || 1)
      });
      const defaults = applyRunDefaultsFromModel(setupModel, run.values);
      if (defaults.changed) {
        const updated = await updateRun(run.runId, { values: defaults.values });
        if (updated) run = updated;
      }
      await load('refresh');
      return run;
    },
    [load, setupModel, templateId, templateReady]
  );

  const deleteRunById = useCallback(
    async (runId: string) => {
      await deleteRunCascade(runId);
      await load('refresh');
    },
    [load]
  );

  const renameRunById = useCallback(
    async (runId: string, title: string) => {
      const updated = await renameRun(runId, title);
      if (updated) await load('refresh');
      return updated;
    },
    [load]
  );

  return {
    initialLoading,
    refreshing,
    error,
    runs,
    runTree,
    templateId,
    templateName,
    templateReady,
    setupModel,
    load,
    createNewRun,
    deleteRunById,
    renameRunById
  };
}
