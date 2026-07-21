import { StyleSheet, View } from 'react-native';

import { spacing } from '../../constants/theme';
import { PrimaryButton } from '../mobile';

type Props = {
  showBack?: boolean;
  onBack?: () => void;
  primaryLabel: string;
  onPrimary: () => void;
  primaryDisabled?: boolean;
  primaryLoading?: boolean;
};

export function WizardFooter({
  showBack,
  onBack,
  primaryLabel,
  onPrimary,
  primaryDisabled,
  primaryLoading
}: Props) {
  return (
    <View style={styles.row}>
      {showBack && onBack ? (
        <PrimaryButton label="Zurück" variant="secondary" onPress={onBack} style={styles.back} />
      ) : null}
      <PrimaryButton
        label={primaryLabel}
        onPress={onPrimary}
        disabled={primaryDisabled}
        loading={primaryLoading}
        style={showBack ? styles.primary : styles.full}
      />
    </View>
  );
}

const styles = StyleSheet.create({
  row: {
    flexDirection: 'row',
    gap: spacing.sm,
    alignItems: 'stretch'
  },
  back: {
    flex: 1
  },
  primary: {
    flex: 2
  },
  full: {
    flex: 1
  }
});
