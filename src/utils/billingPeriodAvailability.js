function toLocalDate(value) {
  if (value instanceof Date) {
    const copy = new Date(value);
    return Number.isNaN(copy.getTime()) ? null : copy;
  }

  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const [year, month, day] = value.split("-").map(Number);
    const parsed = new Date(year, month - 1, day);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  const parsed = value === undefined ? new Date() : new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

export function getBillingAsOfDate(value = new Date()) {
  const parsed = toLocalDate(value);
  if (!parsed) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

export function getLatestNormalBillingMonthStart(asOfDate = new Date()) {
  const asOf = getBillingAsOfDate(asOfDate);
  if (!asOf) return null;

  const monthOffset = asOf.getDate() >= 25 ? 0 : -1;
  return new Date(asOf.getFullYear(), asOf.getMonth() + monthOffset, 1);
}

export function isNormalBillingMonthAvailable(
  billingMonth,
  asOfDate = new Date(),
) {
  const monthDate = toLocalDate(billingMonth);
  const latestAvailable = getLatestNormalBillingMonthStart(asOfDate);
  if (!monthDate || !latestAvailable) return false;

  const monthStart = new Date(
    monthDate.getFullYear(),
    monthDate.getMonth(),
    1,
  );
  return monthStart <= latestAvailable;
}
