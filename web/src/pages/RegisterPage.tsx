import { useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import { useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { AUTH_COPY, registerSchema, useAuth, type RegisterFormValues } from '@buew/shared';

import { AuthShell } from '../components/AuthShell';
import { PasswordField, TextField } from '../components/FormFields';

export function RegisterPage() {
  const { signUp, online } = useAuth();
  const navigate = useNavigate();
  const [formError, setFormError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);

  const {
    register,
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
    setSuccess(null);
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
    setSuccess('Konto erstellt. Falls eine Bestätigung nötig ist, prüfe deine E-Mails.');
    navigate('/', { replace: true });
  });

  return (
    <AuthShell
      title={AUTH_COPY.registerTitle}
      subtitle={AUTH_COPY.registerSubtitle}
      footer={<Link to="/login">Bereits registriert? Anmelden</Link>}
    >
      {!online ? <p className="banner warning">{AUTH_COPY.offline}</p> : null}
      <form className="stack" onSubmit={onSubmit} noValidate>
        <TextField label="Vorname" error={errors.firstName?.message} {...register('firstName')} />
        <TextField label="Nachname" error={errors.lastName?.message} {...register('lastName')} />
        <TextField
          label="E-Mail"
          type="email"
          autoComplete="email"
          error={errors.email?.message}
          {...register('email')}
        />
        <PasswordField
          label="Passwort"
          autoComplete="new-password"
          error={errors.password?.message}
          {...register('password')}
        />
        <PasswordField
          label="Passwort wiederholen"
          autoComplete="new-password"
          error={errors.confirmPassword?.message}
          {...register('confirmPassword')}
        />
        {formError ? <p className="banner danger">{formError}</p> : null}
        {success ? <p className="banner success">{success}</p> : null}
        <button className="primary-btn" type="submit" disabled={isSubmitting || !online}>
          {isSubmitting ? 'Registrieren…' : 'Registrieren'}
        </button>
      </form>
    </AuthShell>
  );
}
