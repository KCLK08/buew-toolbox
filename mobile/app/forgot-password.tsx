import { useState } from 'react';
import { Link } from 'expo-router';
import { Controller, useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import {
  AUTH_COPY,
  forgotPasswordSchema,
  useAuth,
  type ForgotPasswordFormValues,
  colors
} from '@buew/shared';
import { StyleSheet, Text } from 'react-native';

import { AuthScreen } from '../src/components/AuthScreen';
import { TextField } from '../src/components/FormFields';
import { PrimaryButton } from '../src/components/PrimaryButton';

export default function ForgotPasswordScreen() {
  const { resetPassword, online } = useAuth();
  const [formError, setFormError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);

  const {
    control,
    handleSubmit,
    formState: { errors, isSubmitting }
  } = useForm<ForgotPasswordFormValues>({
    resolver: zodResolver(forgotPasswordSchema),
    defaultValues: { email: '' }
  });

  const onSubmit = handleSubmit(async (values) => {
    setFormError(null);
    setSuccess(null);
    const result = await resetPassword(values);
    if (result.error) {
      setFormError(result.error);
      return;
    }
    setSuccess('Falls ein Konto existiert, wurde ein Reset-Link per E-Mail gesendet.');
  });

  return (
    <AuthScreen
      title={AUTH_COPY.forgotTitle}
      subtitle={AUTH_COPY.forgotSubtitle}
      footer={
        <Link href="/login" style={styles.link}>
          Zurück zum Login
        </Link>
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
      {formError ? <Text style={styles.error}>{formError}</Text> : null}
      {success ? <Text style={styles.success}>{success}</Text> : null}
      <PrimaryButton
        label="Reset-Link senden"
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
  },
  success: {
    color: colors.success,
    fontFamily: 'SpaceGrotesk_400Regular'
  }
});
