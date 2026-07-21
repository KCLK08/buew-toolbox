import type { ReactNode } from 'react';

type AuthShellProps = {
  title: string;
  subtitle: string;
  children: ReactNode;
  footer?: ReactNode;
};

export function AuthShell({ title, subtitle, children, footer }: AuthShellProps) {
  return (
    <div className="auth-page">
      <div className="auth-card">
        <p className="brand">BÜW-Toolbox</p>
        <h1>{title}</h1>
        <p className="sub">{subtitle}</p>
        {children}
        {footer ? <div className="auth-footer">{footer}</div> : null}
      </div>
    </div>
  );
}
