import { useState, type InputHTMLAttributes } from 'react';

type PasswordFieldProps = InputHTMLAttributes<HTMLInputElement> & {
  label: string;
  error?: string;
};

export function PasswordField({ label, error, id, ...props }: PasswordFieldProps) {
  const [visible, setVisible] = useState(false);
  const fieldId = id ?? props.name;

  return (
    <label className="field" htmlFor={fieldId}>
      <span>{label}</span>
      <div className="password-wrap">
        <input
          {...props}
          id={fieldId}
          type={visible ? 'text' : 'password'}
          autoComplete={props.autoComplete ?? 'current-password'}
        />
        <button
          type="button"
          className="ghost-btn"
          onClick={() => setVisible((v) => !v)}
          aria-label={visible ? 'Passwort verbergen' : 'Passwort anzeigen'}
        >
          {visible ? 'Verbergen' : 'Anzeigen'}
        </button>
      </div>
      {error ? <em className="error">{error}</em> : null}
    </label>
  );
}

type TextFieldProps = InputHTMLAttributes<HTMLInputElement> & {
  label: string;
  error?: string;
};

export function TextField({ label, error, id, ...props }: TextFieldProps) {
  const fieldId = id ?? props.name;
  return (
    <label className="field" htmlFor={fieldId}>
      <span>{label}</span>
      <input {...props} id={fieldId} />
      {error ? <em className="error">{error}</em> : null}
    </label>
  );
}
