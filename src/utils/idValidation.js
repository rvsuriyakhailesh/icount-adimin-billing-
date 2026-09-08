const ID_PATTERN = /^[A-Za-z0-9]{1,8}$/;

export function normalizeIdInput(value, { allowStandaloneBlank = false } = {}) {
  const normalizedValue = String(value ?? "").trim().toUpperCase();

  if (!normalizedValue) {
    return "";
  }

  if (
    allowStandaloneBlank &&
    ["STANDALONE", "SINGLE"].includes(normalizedValue)
  ) {
    return "";
  }

  return normalizedValue;
}

export function isValidIdInput(value) {
  return ID_PATTERN.test(String(value ?? "").trim());
}

export function isStandaloneComplexCode(value) {
  const normalizedValue = String(value ?? "").trim().toUpperCase();
  return !normalizedValue || ["STANDALONE", "SINGLE"].includes(normalizedValue);
}

export function getDisplayComplexCode(value) {
  if (isStandaloneComplexCode(value)) {
    return "-";
  }

  return normalizeIdInput(value);
}

export function getIdValidationMessage(label) {
  return `${label} must be alphanumeric and contain up to 8 characters.`;
}
