import { useState } from 'react';
import { Link } from 'react-router-dom';
import { useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import {
  AUTH_COPY,
  forgotPasswordSchema,
  useAuth,
  type ForgotPasswordFormValues
} from '@buew/shared';

import { AuthShell } from '../components/AuthShell';
import { TextField } from '../components/FormFields';

export function ForgotPasswordPage() {
  const { resetPassword, online } = useAuth();
  const [formError, setFormError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);

  const {
    register,
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
    <AuthShell
      title={AUTH_COPY.forgotTitle}
      subtitle={AUTH_COPY.forgotSubtitle}
      footer={<Link to="/login">Zurück zum Login</Link>}
    >
      {!online ? <p className="banner warning">{AUTH_COPY.offline}</p> : null}
      <form className="stack" onSubmit={onSubmit} noValidate>
        <TextField
          label="E-Mail"
          type="email"
          autoComplete="email"
          error={errors.email?.message}
          {...register('email')}
        />
        {formError ? <p className="banner danger">{formError}</p> : null}
        {success ? <p className="banner success">{success}</p> : null}
        <button className="primary-btn" type="submit" disabled={isSubmitting || !online}>
          {isSubmitting ? 'Senden…' : 'Reset-Link senden'}
        </button>
      </form>
    </AuthShell>
  );
}
