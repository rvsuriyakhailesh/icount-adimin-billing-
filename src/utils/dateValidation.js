export function isValidFourDigitYear(dateValue) {
  if (!dateValue) {
    return true;
  }

  const year = String(dateValue).split("-")[0];
  return /^\d{4}$/.test(year);
}

export function isValidDateInputValue(dateValue) {
  if (!dateValue) {
    return true;
  }

  return /^\d{4}-\d{2}-\d{2}$/.test(String(dateValue)) && isValidFourDigitYear(dateValue);
}

export function isValidDateValue(dateValue) {
  if (!isValidDateInputValue(dateValue)) {
    return false;
  }

  const [year, month, day] = String(dateValue).split("-").map(Number);
  const parsed = new Date(Date.UTC(year, month - 1, day));

  return (
    parsed.getUTCFullYear() === year &&
    parsed.getUTCMonth() === month - 1 &&
    parsed.getUTCDate() === day
  );
}
