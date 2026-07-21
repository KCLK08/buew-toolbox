import { useState } from 'react';
import { Link, useRouter } from 'expo-router';
import { Controller, useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { AUTH_COPY, registerSchema, useAuth, type RegisterFormValues, colors } from '@buew/shared';
import { StyleSheet, Text } from 'react-native';

import { AuthScreen } from '../src/components/AuthScreen';
import { PasswordField, TextField } from '../src/components/FormFields';
import { PrimaryButton } from '../src/components/PrimaryButton';
import { useBiometric } from '../src/providers/BiometricProvider';

export default function RegisterScreen() {
  const { signUp, online } = useAuth();
  const biometric = useBiometric();
  const router = useRouter();
  const [formError, setFormError] = useState<string | null>(null);

  const {
    control,
    handleSubmit,
    formState: { errors, isSubmitting }
  } = useForm<RegisterFormValues>({
    resolver: zodResolver(registerSchema),
    defaultValues: {
      firstName: '',
      lastName: '',
      email: '',
      password: '',
      confirmPassword: ''
    }
  });

  const onSubmit = handleSubmit(async (values) => {
    setFormError(null);
    const result = await signUp({
      firstName: values.firstName,
      lastName: values.lastName,
      email: values.email,
      password: values.password
    });
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
      title={AUTH_COPY.registerTitle}
      subtitle={AUTH_COPY.registerSubtitle}
      footer={
        <Link href="/login" style={styles.link}>
          Bereits registriert? Anmelden
        </Link>
      }
    >
      {!online ? <Text style={styles.warning}>{AUTH_COPY.offline}</Text> : null}
      <Controller
        control={control}
        name="firstName"
        render={({ field: { onChange, onBlur, value } }) => (
          <TextField
            label="Vorname"
            value={value}
            onBlur={onBlur}
            onChangeText={onChange}
            error={errors.firstName?.message}
          />
        )}
      />
      <Controller
        control={control}
        name="lastName"
        render={({ field: { onChange, onBlur, value } }) => (
          <TextField
            label="Nachname"
            value={value}
            onBlur={onBlur}
            onChangeText={onChange}
            error={errors.lastName?.message}
          />
        )}
      />
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
            autoComplete="new-password"
            value={value}
            onBlur={onBlur}
            onChangeText={onChange}
            error={errors.password?.message}
          />
        )}
      />
      <Controller
        control={control}
        name="confirmPassword"
        render={({ field: { onChange, onBlur, value } }) => (
          <PasswordField
            label="Passwort wiederholen"
            autoComplete="new-password"
            value={value}
            onBlur={onBlur}
            onChangeText={onChange}
            error={errors.confirmPassword?.message}
          />
        )}
      />
      {formError ? <Text style={styles.error}>{formError}</Text> : null}
      <PrimaryButton
        label="Registrieren"
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
