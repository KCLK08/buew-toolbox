import { Redirect } from 'expo-router';
import { useAuth } from '@buew/shared';

import { ToolWebScreen } from '../src/components/ToolWebScreen';
import { useBiometric } from '../src/providers/BiometricProvider';

export default function BautagebuchScreen() {
  const { isAuthenticated } = useAuth();
  const biometric = useBiometric();

  if (!isAuthenticated) return <Redirect href="/login" />;
  if (biometric.needsUnlock) return <Redirect href="/biometric-unlock" />;

  return <ToolWebScreen title="Bautagebuch" webPath="/bautagebuch/" />;
}
