import { z } from 'zod';

export const passwordSchema = z
  .string()
  .min(8, 'Mindestens 8 Zeichen')
  .regex(/[A-Z]/, 'Mindestens ein Großbuchstabe')
  .regex(/[a-z]/, 'Mindestens ein Kleinbuchstabe')
  .regex(/[0-9]/, 'Mindestens eine Zahl')
  .regex(/[^A-Za-z0-9]/, 'Mindestens ein Sonderzeichen');

export const loginSchema = z.object({
  email: z.string().trim().email('Ungültige E-Mail-Adresse'),
  password: z.string().min(1, 'Passwort ist erforderlich')
});

export const registerSchema = z
  .object({
    firstName: z.string().trim().min(1, 'Vorname ist erforderlich').max(80),
    lastName: z.string().trim().min(1, 'Nachname ist erforderlich').max(80),
    email: z.string().trim().email('Ungültige E-Mail-Adresse'),
    password: passwordSchema,
    confirmPassword: z.string().min(1, 'Passwortwiederholung ist erforderlich')
  })
  .refine((data) => data.password === data.confirmPassword, {
    message: 'Passwörter stimmen nicht überein',
    path: ['confirmPassword']
  });

export const forgotPasswordSchema = z.object({
  email: z.string().trim().email('Ungültige E-Mail-Adresse')
});

export type LoginFormValues = z.infer<typeof loginSchema>;
export type RegisterFormValues = z.infer<typeof registerSchema>;
export type ForgotPasswordFormValues = z.infer<typeof forgotPasswordSchema>;
