import { useState } from 'react';
import { Link, Redirect, useRouter } from 'expo-router';
import { Controller, useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { AUTH_COPY, loginSchema, useAuth, type LoginFormValues, colors } from '@buew/shared';
import { StyleSheet, Text } from 'react-native';

import { AuthScreen } from '../src/components/AuthScreen';
import { PasswordField, TextField } from '../src/components/FormFields';
import { PrimaryButton } from '../src/components/PrimaryButton';
import { useBiometric } from '../src/providers/BiometricProvider';

export default function LoginScreen() {
  const { signIn, online, isAuthenticated } = useAuth();
  const biometric = useBiometric();
  const router = useRouter();
  const [formError, setFormError] = useState<string | null>(null);

  const {
    control,
    handleSubmit,
    formState: { errors, isSubmitting }
  } = useForm<LoginFormValues>({
    resolver: zodResolver(loginSchema),
    defaultValues: { email: '', password: '' }
  });

  if (isAuthenticated && !biometric.needsUnlock) {
    return <Redirect href="/dashboard" />;
  }

  const onSubmit = handleSubmit(async (values) => {
    setFormError(null);
    const result = await signIn(values);
    if (result.error) {
      setFormError(result.error);
      return;
    }
    biometric.markUnlocked();
    biometric.promptEnableAfterLogin();
    router.replace('/dashboard');
  });

  return (
    <AuthScreen
      title={AUTH_COPY.loginTitle}
      subtitle={AUTH_COPY.loginSubtitle}
      footer={
        <>
          <Link href="/register" style={styles.link}>
            Konto erstellen
          </Link>
          <Link href="/forgot-password" style={styles.link}>
            Passwort vergessen
          </Link>
        </>
      }
    >
      {!online ? <Text style={styles.warning}>{AUTH_COPY.offline}</Text> : null}
      <Controller
        control={control}
        name="email"
        render={({ field: { onChange, onBlur, value } }) => (
          <TextField
            label="E-Mail"
            autoCapitalize="none"
            keyboardType="email-address"
            autoComplete="email"
            value={value}
            onBlur={onBlur}
            onChangeText={onChange}
            error={errors.email?.message}
          />
        )}
      />
      <Controller
        control={control}
        name="password"
        render={({ field: { onChange, onBlur, value } }) => (
          <PasswordField
            label="Passwort"
            autoComplete="password"
            value={value}
            onBlur={onBlur}
            onChangeText={onChange}
            error={errors.password?.message}
          />
        )}
      />
      {formError ? <Text style={styles.error}>{formError}</Text> : null}
      <PrimaryButton
        label="Login"
        loading={isSubmitting}
        disabled={!online}
        onPress={() => void onSubmit()}
      />
    </AuthScreen>
  );
}

const styles = StyleSheet.create({
  link: {
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 14
  },
  warning: {
    color: colors.accent2,
    backgroundColor: 'rgba(211,84,60,0.12)',
    padding: 10,
    borderRadius: 12,
    fontFamily: 'SpaceGrotesk_400Regular'
  },
  error: {
    color: colors.danger,
    fontFamily: 'SpaceGrotesk_400Regular'
  }
});
