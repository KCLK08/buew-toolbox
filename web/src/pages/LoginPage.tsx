import { useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import { useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { AUTH_COPY, loginSchema, useAuth, type LoginFormValues } from '@buew/shared';

import { AuthShell } from '../components/AuthShell';
import { PasswordField, TextField } from '../components/FormFields';

export function LoginPage() {
  const { signIn, online } = useAuth();
  const navigate = useNavigate();
  const [formError, setFormError] = useState<string | null>(null);

  const {
    register,
    handleSubmit,
    formState: { errors, isSubmitting }
  } = useForm<LoginFormValues>({
    resolver: zodResolver(loginSchema),
    defaultValues: { email: '', password: '' }
  });

  const onSubmit = handleSubmit(async (values) => {
    setFormError(null);
    const result = await signIn(values);
    if (result.error) {
      setFormError(result.error);
      return;
    }
    navigate('/', { replace: true });
  });

  return (
    <AuthShell
      title={AUTH_COPY.loginTitle}
      subtitle={AUTH_COPY.loginSubtitle}
      footer={
        <>
          <Link to="/register">Konto erstellen</Link>
          <Link to="/forgot-password">Passwort vergessen</Link>
        </>
      }
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
        <PasswordField
          label="Passwort"
          autoComplete="current-password"
          error={errors.password?.message}
          {...register('password')}
        />
        {formError ? <p className="banner danger">{formError}</p> : null}
        <button className="primary-btn" type="submit" disabled={isSubmitting || !online}>
          {isSubmitting ? 'Anmelden…' : 'Login'}
        </button>
      </form>
    </AuthShell>
  );
}
