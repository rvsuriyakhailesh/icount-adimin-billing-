export function preventNumberWheel(event) {
  event.currentTarget.blur();
}

export function sanitizeAmount(value) {
  const cleaned = String(value ?? "").replace(/[^\d.]/g, "");
  const parts = cleaned.split(".");

  return parts.length > 1
    ? `${parts[0]}.${parts.slice(1).join("")}`
    : cleaned;
}
