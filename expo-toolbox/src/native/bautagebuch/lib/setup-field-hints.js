const SHIFT_FIELD_NAMES = new Set(['Check Box1', 'Check Box2', 'Check Box3']);
const GEWERK_FIELD_NAMES = new Set(['Text3', 'Text5', 'Text6', 'Text7', 'Text8']);

export function resolveSetupFieldType(field, detectedFields = []) {
  const explicit = String(field?.type || '').trim();
  if (explicit === 'select') return 'dropdown';
  if (explicit) return explicit;
  const fieldId = String(field?.fieldId || '').trim();
  const detected = (detectedFields || []).find((entry) => String(entry?.fieldId || '') === fieldId);
  return String(detected?.type || 'text');
}

export function checkboxBehaviorHint(fieldName) {
  const name = String(fieldName || '').trim();
  if (SHIFT_FIELD_NAMES.has(name)) {
    return 'Schicht-Auswahl: Es darf nur eine Option (Früh-, Spät- oder Nachtschicht) gleichzeitig aktiv sein.';
  }
  if (GEWERK_FIELD_NAMES.has(name)) {
    return 'Gewerk-Felder: Mindestens eines soll ausgefüllt werden; mehrere gleichzeitig sind erlaubt.';
  }
  return 'Unabhängige Checkbox — kann ohne Einfluss auf andere Checkboxen gesetzt werden.';
}

export function isCheckboxField(field, detectedFields = []) {
  return resolveSetupFieldType(field, detectedFields) === 'checkbox';
}

export function readCheckboxDefault(field) {
  const raw = field?.defaultValue;
  if (raw === true || raw === 'true') return true;
  if (raw === false || raw === 'false') return false;
  return false;
}

export function writeCheckboxDefault(checked) {
  return checked ? 'true' : 'false';
}
