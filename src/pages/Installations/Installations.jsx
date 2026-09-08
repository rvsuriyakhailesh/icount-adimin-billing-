import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import "./Installations.css";
import {
  isValidDateInputValue,
  isValidDateValue,
} from "../../utils/dateValidation";
import {
  getDisplayComplexCode,
  normalizeIdInput,
} from "../../utils/idValidation";
import { preventNumberWheel, sanitizeAmount } from "../../utils/numberInput";

const GUIDED_SCROLL_DURATION_MS = 800;
const GUIDED_SCROLL_TOP_OFFSET_PX = 24;
const BILLING_API_BASE =
  import.meta.env.VITE_BILLING_API_BASE || "http://localhost:4100/api";

function getCanonicalExpenseId(expense, screenCode = "") {
  const explicitId = normalizeValue(expense?.expenseId);
  if (explicitId) return explicitId;

  return [
    normalizeIdInput(expense?.screenCode) || normalizeIdInput(screenCode),
    normalizeValue(expense?.expenseType).toLowerCase(),
    normalizeValue(expense?.expenseDate),
    Number(expense?.amount || 0),
    normalizeValue(expense?.comments),
    normalizeValue(expense?.createdAt || expense?.updatedAt),
  ].join("::");
}

function dedupeExpenses(expenses, screenCode = "") {
  const seen = new Set();

  return (Array.isArray(expenses) ? expenses : []).filter((expense) => {
    const expenseId = getCanonicalExpenseId(expense, screenCode);
    if (seen.has(expenseId)) return false;
    seen.add(expenseId);
    return true;
  });
}

function getOwnedExpenses(expenses, screenCode = "", screenName = "") {
  const canonicalScreenCode = normalizeIdInput(screenCode);

  return dedupeExpenses(expenses, canonicalScreenCode)
    .filter((expense) => {
      const expenseScreenCode = normalizeIdInput(expense?.screenCode);
      return !expenseScreenCode || expenseScreenCode === canonicalScreenCode;
    })
    .map((expense) => ({
      ...expense,
      expenseId: getCanonicalExpenseId(expense, canonicalScreenCode),
      screenCode: canonicalScreenCode,
      screenName: normalizeValue(expense?.screenName) || screenName,
    }));
}

function buildStage2PersistenceSnapshot(record) {
  return {
    dateOfDispatch: record?.dateOfDispatch || "",
    installationDate: record?.installationDate || "",
    liveDate: record?.liveDate || "",
    trialPeriod: record?.trialPeriod || "",
    trialPeriodExtension: getNormalizedTrialPeriodExtension(record),
    totalTrialPeriodExtension: getNormalizedTrialPeriodExtension(record),
    trialExtension: getNormalizedTrialPeriodExtension(record),
    billingStartDate: record?.billingStartDate || "",
    installationStageStatus: record?.installationStageStatus || "",
    installationBlocker: record?.installationBlocker || "",
    blockerReason: record?.blockerReason || "",
    remarks: record?.remarks || "",
    extensionRemarks: record?.extensionRemarks || "",
    installationDetailsCompleted: record?.installationDetailsCompleted === true,
    saved2B: record?.saved2B === true,
    saved2BAt: record?.saved2BAt || "",
    otherInstallationExpensesApplicable:
      record?.otherInstallationExpensesApplicable || "",
    installationExpenses: getOwnedExpenses(
      record?.installationExpenses,
      record?.screenCode,
      record?.screenName,
    ),
    installationExpensesTotal: Number(record?.installationExpensesTotal || 0),
    complexInstallationExpensesTotal: Number(
      record?.complexInstallationExpensesTotal || 0,
    ),
    installationExpensesSavedAt: record?.installationExpensesSavedAt || "",
    expenseCycleStartedAt: record?.expenseCycleStartedAt || "",
  };
}

async function resolveBackendSiteId(record) {
  if (record?.backendSiteId) return record.backendSiteId;

  const screenCode = normalizeIdInput(record?.screenCode);
  if (!screenCode) return "";

  const response = await fetch(`${BILLING_API_BASE}/sites`);
  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(body?.message || body?.error || "Unable to locate backend site.");
  }

  const sites = Array.isArray(body?.data) ? body.data : [];
  const match = sites.find(
    (site) => normalizeIdInput(site?.siteId) === screenCode,
  );
  return match?.id || "";
}

async function persistStage2Record(record) {
  const backendSiteId = await resolveBackendSiteId(record);
  if (!backendSiteId) {
    throw new Error(`Backend site not found for ${record?.screenCode || "selected screen"}.`);
  }

  const response = await fetch(`${BILLING_API_BASE}/sites/${backendSiteId}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      stage2Data: buildStage2PersistenceSnapshot(record),
    }),
  });
  const body = await response.json().catch(() => ({}));

  if (!response.ok) {
    const detail = body?.details?.fieldErrors
      ? Object.entries(body.details.fieldErrors)
          .flatMap(([field, messages]) =>
            (Array.isArray(messages) ? messages : [messages]).map(
              (message) => `${field}: ${message}`,
            ),
          )
          .join("; ")
      : "";
    throw new Error(
      [body?.message || body?.error || "Stage 2 could not be saved to the backend.", detail]
        .filter(Boolean)
        .join(" - "),
    );
  }

  return { backendSiteId, site: body?.data || null };
}


function compareColumnValues(leftValue, rightValue) {
  const leftText = leftValue === null || leftValue === undefined ? "" : String(leftValue).trim();
  const rightText = rightValue === null || rightValue === undefined ? "" : String(rightValue).trim();

  const leftDate = /^\d{4}-\d{2}-\d{2}/.test(leftText) ? Date.parse(leftText) : NaN;
  const rightDate = /^\d{4}-\d{2}-\d{2}/.test(rightText) ? Date.parse(rightText) : NaN;
  if (!Number.isNaN(leftDate) && !Number.isNaN(rightDate)) return leftDate - rightDate;

  const leftNumber = Number(leftText.replace(/[^0-9.-]/g, ""));
  const rightNumber = Number(rightText.replace(/[^0-9.-]/g, ""));
  if (leftText && rightText && Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) {
    return leftNumber - rightNumber;
  }

  return leftText.localeCompare(rightText, undefined, { numeric: true, sensitivity: "base" });
}

function SortableHeader({ label, sortKey, activeKey, direction, onSort, style, className }) {
  const active = activeKey === sortKey;
  return (
    <th className={className} style={style}>
      <button
        type="button"
        onClick={() => onSort(sortKey)}
        style={{
          appearance: "none",
          border: 0,
          background: "transparent",
          padding: 0,
          margin: 0,
          font: "inherit",
          fontWeight: "inherit",
          color: "inherit",
          cursor: "pointer",
          textAlign: "inherit",
        }}
        title={`Sort ${label}`}
      >
        {label} {active ? (direction === "asc" ? "▲" : "▼") : ""}
      </button>
    </th>
  );
}

function guidedScrollToElement(
  element,
  {
    duration = GUIDED_SCROLL_DURATION_MS,
    topOffset = GUIDED_SCROLL_TOP_OFFSET_PX,
  } = {},
) {
  if (!element) {
    return;
  }

  const startY = window.scrollY;
  const targetY = Math.max(
    0,
    element.getBoundingClientRect().top + window.scrollY - topOffset,
  );
  const distance = targetY - startY;

  if (Math.abs(distance) < 2) {
    return;
  }

  const startTime = performance.now();

  function easeInOutCubic(progress) {
    return progress < 0.5
      ? 4 * progress * progress * progress
      : 1 - Math.pow(-2 * progress + 2, 3) / 2;
  }

  function step(now) {
    const elapsed = now - startTime;
    const progress = Math.min(elapsed / duration, 1);
    const easedProgress = easeInOutCubic(progress);

    window.scrollTo(0, startY + distance * easedProgress);

    if (progress < 1) {
      window.requestAnimationFrame(step);
    }
  }

  window.requestAnimationFrame(step);
}

const blockerReasonOptions = [
  "Site Not Ready",
  "Customer / Client Dependency",
  "Material / Equipment Issue",
  "Technical Issue",
  "Manpower / Labour Issue",
  "Permission / Approval Pending",
  "Access / Location Issue",
  "Power / Infrastructure Issue",
  "Internal Coordination",
  "Commercial Issue",
  "Other",
];

const downloadStageOptions = [
  "All",
  "Transport",
  "Installation Pending",
  "Action Required",
  "Installation Blocker",
  "Completed",
  "Good to Go",
];

const readinessStageFilterOptions = [
  { value: "all", label: "All" },
  { value: "ready", label: "Ready for Stage 3" },
  { value: "future", label: "Future Billing" },
  { value: "partial", label: "Partially Completed" },
  { value: "overdue", label: "Overdue" },
  { value: "hold", label: "On Hold" },
  { value: "installation-pending", label: "Installation Pending" },
];

function normalizeValue(value) {
  return String(value || "").trim();
}

function getDisplayValue(value) {
  return normalizeValue(value) || "-";
}

function isBillingCommercialLocked(record) {
  return (
    record?.billingCommercialLocked === true ||
    ["Submitted to Billing Team", "Sent to Billing Team"].includes(
      normalizeValue(record?.billingVerificationStatus),
    ) ||
    Boolean(normalizeValue(record?.billingCommercialLockedAt))
  );
}

const billingLockedInstallationFields = new Set([
  "installationDate",
  "liveDate",
  "trialPeriod",
  "trialPeriodExtension",
  "totalTrialPeriodExtension",
  "trialExtension",
  "billingStartDate",
  "installationStageStatus",
  "installationBlocker",
  "blockerReason",
]);

function formatExpenseAmount(value) {
  return Number(value).toLocaleString("en-IN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function isValidExpenseAmount(value) {
  return /^\d+(\.\d+)?$/.test(value) && Number(value) > 0;
}

function getDefaultExpenseState() {
  return {
    applicable: "",
    expenses: [],
  };
}

function getNormalizedTrialPeriodExtension(record) {
  return normalizeValue(
    record?.trialPeriodExtension ??
      record?.totalTrialPeriodExtension ??
      record?.trialExtension,
  );
}

function getNormalizedRemarks(record) {
  return normalizeValue(record?.remarks ?? record?.extensionRemarks);
}

function getRemarksInputValue(record) {
  return String(record?.remarks ?? record?.extensionRemarks ?? "");
}

function isInstallationDetailsCompleted(record) {
  if (record?.installationDetailsCompleted !== true) {
    return false;
  }

  const dateOfDispatch = normalizeValue(record?.dateOfDispatch);
  const installationDate = normalizeValue(record?.installationDate);
  const liveDate = normalizeValue(record?.liveDate);
  const billingStartDate = normalizeValue(record?.billingStartDate);
  const trialPeriod = normalizeValue(record?.trialPeriod);
  const trialPeriodExtension = getNormalizedTrialPeriodExtension(record);

  if (
    !dateOfDispatch ||
    !installationDate ||
    !liveDate ||
    !billingStartDate
  ) {
    return false;
  }

  if (
    !isValidCalendarDateValue(dateOfDispatch) ||
    !isValidCalendarDateValue(installationDate) ||
    !isValidCalendarDateValue(liveDate) ||
    !isValidCalendarDateValue(billingStartDate)
  ) {
    return false;
  }

  if (
    !isValidOptionalWholeNumber(trialPeriod) ||
    !isValidOptionalWholeNumber(trialPeriodExtension)
  ) {
    return false;
  }

  if (isFutureDateValue(installationDate) || isFutureDateValue(liveDate)) {
    return false;
  }

  const parsedDispatchDate = parseDate(dateOfDispatch);
  const parsedInstallationDate = parseDate(installationDate);
  const parsedLiveDate = parseDate(liveDate);
  const parsedBillingStartDate = parseDate(billingStartDate);

  if (
    !parsedDispatchDate ||
    !parsedInstallationDate ||
    !parsedLiveDate ||
    !parsedBillingStartDate
  ) {
    return false;
  }

  if (parsedDispatchDate > parsedInstallationDate) {
    return false;
  }

  if (parsedInstallationDate > parsedLiveDate) {
    return false;
  }

  if (parsedDispatchDate > parsedLiveDate) {
    return false;
  }

  if (parsedLiveDate > parsedBillingStartDate) {
    return false;
  }

  return true;
}

function formatDateValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

function parseDate(value) {
  if (!isValidCalendarDateValue(value)) {
    return null;
  }

  const [year, month, day] = String(value).split("-").map(Number);
  const parsed = new Date(Date.UTC(year, month - 1, day));
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function parseNonNegativeWholeNumber(value) {
  const normalizedValue = normalizeValue(value);

  if (!normalizedValue) {
    return 0;
  }

  if (!/^\d+$/.test(normalizedValue)) {
    return null;
  }

  return Number(normalizedValue);
}

function isValidOptionalWholeNumber(value) {
  const normalizedValue = normalizeValue(value);
  return !normalizedValue || /^\d+$/.test(normalizedValue);
}

function isValidCalendarDateValue(dateValue) {
  const normalizedValue = normalizeValue(dateValue);

  if (!normalizedValue) {
    return false;
  }

  if (!isValidDateInputValue(normalizedValue) || !isValidDateValue(normalizedValue)) {
    return false;
  }

  const [year, month, day] = normalizedValue.split("-").map(Number);
  const parsedDate = new Date(Date.UTC(year, month - 1, day));

  return (
    parsedDate.getUTCFullYear() === year &&
    parsedDate.getUTCMonth() === month - 1 &&
    parsedDate.getUTCDate() === day
  );
}

function isFutureDateValue(dateValue) {
  const normalizedValue = normalizeValue(dateValue);

  if (!isValidDateInputValue(normalizedValue) || !isValidDateValue(normalizedValue)) {
    return false;
  }

  const [year, month, day] = normalizedValue.split("-").map(Number);
  const today = new Date();

  const yearDiff = year - today.getFullYear();
  if (yearDiff !== 0) {
    return yearDiff > 0;
  }

  const monthDiff = month - 1 - today.getMonth();
  if (monthDiff !== 0) {
    return monthDiff > 0;
  }

  return day > today.getDate();
}

function calculateBillingStartDate(
  liveDate,
  trialPeriod,
  trialPeriodExtension,
) {
  const parsedLiveDate = parseDate(liveDate);

  if (!parsedLiveDate) {
    return "";
  }

  const trialDays = parseNonNegativeWholeNumber(trialPeriod);
  const extensionDays = parseNonNegativeWholeNumber(trialPeriodExtension);

  if (trialDays === null || extensionDays === null) {
    return "";
  }

  const billingStartDate = new Date(parsedLiveDate);
  billingStartDate.setDate(
    billingStartDate.getDate() + trialDays + extensionDays + 1,
  );

  return formatDateValue(billingStartDate);
}

function calculateInstallationBillingDate(record) {
  const installationDate = normalizeValue(record.installationDate);
  const liveDate = normalizeValue(record.liveDate);
  const trialPeriod = normalizeValue(record.trialPeriod);
  const additionalTrialPeriod = getNormalizedTrialPeriodExtension(record);

  if (!installationDate || !liveDate) {
    return "";
  }

  const parsedLiveDate = parseDate(liveDate);
  const trialDays = parseNonNegativeWholeNumber(trialPeriod);
  const additionalDays = parseNonNegativeWholeNumber(additionalTrialPeriod);

  if (!parsedLiveDate || trialDays === null || additionalDays === null) {
    return "";
  }

  const billingDate = new Date(parsedLiveDate);
  billingDate.setDate(billingDate.getDate() + trialDays + additionalDays);

  return formatDateValue(billingDate);
}

function getDaysSinceDispatch(dateOfDispatch) {
  const parsedDispatchDate = parseDate(dateOfDispatch);

  if (!parsedDispatchDate) {
    return null;
  }

  const today = getTodayStart();
  const diffDays = Math.floor((today - parsedDispatchDate) / (1000 * 60 * 60 * 24)) + 1;

  return diffDays;
}

function isBillingWindowOpen(record) {
  if (!isInstallationDetailsCompleted(record)) {
    return false;
  }

  const billingStartDate = normalizeValue(record.billingStartDate);

  if (!billingStartDate) {
    return false;
  }

  const parsedBillingStartDate = parseDate(billingStartDate);
  if (!parsedBillingStartDate) {
    return false;
  }

  const goLiveDate = new Date(parsedBillingStartDate);
  goLiveDate.setDate(goLiveDate.getDate() - 3);

  return getTodayStart() >= goLiveDate;
}

function getBillingMonthStartDate(billingStartDate) {
  const parsedBillingStartDate = parseDate(billingStartDate);

  if (!parsedBillingStartDate) {
    return null;
  }

  return new Date(
    parsedBillingStartDate.getFullYear(),
    parsedBillingStartDate.getMonth(),
    1,
  );
}

function getBillingTreatment(billingStartDate) {
  const parsedBillingStartDate = parseDate(billingStartDate);

  if (!parsedBillingStartDate) {
    return "";
  }

  return parsedBillingStartDate.getDate() === 1 ? "Full Month" : "Pro-rata";
}

function getBillingReadinessStatus(billingStartDate, currentStageStatus = "") {
  const billingMonthStartDate = getBillingMonthStartDate(billingStartDate);

  if (!billingMonthStartDate) {
    return "";
  }

  const today = getTodayStart();
  const currentMonthStartDate = new Date(
    today.getFullYear(),
    today.getMonth(),
    1,
  );
  const normalizedStageStatus = normalizeValue(currentStageStatus).toLowerCase();

  if (normalizedStageStatus.includes("hold")) {
    return "On Hold";
  }

  if (normalizedStageStatus.includes("partial")) {
    return "Partially Completed";
  }

  if (
    normalizedStageStatus.includes("installation pending") ||
    normalizedStageStatus === "transport"
  ) {
    return "Installation Pending";
  }

  if (billingMonthStartDate > currentMonthStartDate) {
    return "Future Billing";
  }

  if (billingMonthStartDate < currentMonthStartDate) {
    return "Overdue";
  }

  return "Ready for Stage 3";
}

function getReadinessStatusTone(status) {
  if (
    status === "Overdue" ||
    status === "Overdue (Ready for Stage 3)" ||
    status === "Installation Blocker"
  ) {
    return "installations-page__readiness-status-badge--warning";
  }

  if (status === "Future Billing") {
    return "installations-page__readiness-status-badge--future";
  }

  if (status === "Ready for Stage 3") {
    return "installations-page__readiness-status-badge--ready";
  }

  if (status === "Partially Completed") {
    return "installations-page__readiness-status-badge--partial";
  }

  if (status === "On Hold") {
    return "installations-page__readiness-status-badge--hold";
  }

  if (status === "Installation Pending") {
    return "installations-page__readiness-status-badge--installation-pending";
  }

  return "installations-page__readiness-status-badge--neutral";
}

function formatReadinessBadgeLabel(status) {
  return status;
}

function formatEligibilityLabel(isEligible) {
  return isEligible ? "Eligible" : "Not Eligible";
}

function InfoTip({ label, tooltip }) {
  return (
    <span className="installations-page__info-wrap">
      <button
        className="installations-page__info-btn"
        type="button"
        aria-label={label}
      >
        i
      </button>
      <span className="installations-page__info-tooltip">{tooltip}</span>
    </span>
  );
}

function normalizeStageFilterValue(value) {
  return normalizeValue(value).toLowerCase();
}

function getReadinessStageKey(record) {
  const readinessStatus = normalizeStageFilterValue(
    record.billingReadinessStatus || record.currentStageStatus,
  );
  const blockerReason = normalizeStageFilterValue(record.blockerReason);

  if (readinessStatus.includes("future billing")) {
    return "future";
  }

  if (
    readinessStatus.includes("ready for stage 3") ||
    readinessStatus.includes("good to go")
  ) {
    return "ready";
  }

  if (
    readinessStatus.includes("installation pending") ||
    readinessStatus === "transport"
  ) {
    return "installation-pending";
  }

  if (
    readinessStatus.includes("overdue") ||
    readinessStatus.includes("action required") ||
    readinessStatus.includes("installation blocker")
  ) {
    return "overdue";
  }

  if (readinessStatus.includes("hold") || blockerReason.includes("hold")) {
    return "hold";
  }

  if (
    readinessStatus.includes("partial") ||
    readinessStatus.includes("completed")
  ) {
    return "partial";
  }

  return "ready";
}

function resolveStage2Status(record) {
  if (isInstallationDetailsCompleted(record)) {
    return isBillingWindowOpen(record) ? "Good to Go" : "Completed";
  }

  const daysSinceDispatch = getDaysSinceDispatch(record.dateOfDispatch);

  if (daysSinceDispatch === null || daysSinceDispatch <= 0) {
    return "Transport";
  }

  if (daysSinceDispatch <= 8) {
    return "Transport";
  }

  if (daysSinceDispatch <= 15) {
    return "Installation Pending";
  }

  if (daysSinceDispatch <= 17) {
    return "Action Required";
  }

  return "Installation Blocker";
}

function getInstallationTableStatusValue(record) {
  const explicitStatus = normalizeValue(record?.installationStageStatus)
    .toLowerCase()
    .trim();
  const hasInstallationDate = Boolean(normalizeValue(record?.installationDate));
  const hasLiveDate = Boolean(normalizeValue(record?.liveDate));
  const hasBillingDate = Boolean(normalizeValue(record?.billingStartDate));

  if (hasInstallationDate && hasLiveDate && hasBillingDate) {
    return "installed-completed";
  }

  if (explicitStatus === "partial-installation" || explicitStatus === "partially-installed") {
    return "partial-installation";
  }

  if (explicitStatus === "installation-blocker" || explicitStatus === "installation-pending") {
    return "installation-blocker";
  }

  if (explicitStatus === "installed-completed") {
    return hasInstallationDate ? "partial-installation" : "installation-blocker";
  }

  if (hasInstallationDate) {
    return "partial-installation";
  }

  return "installation-blocker";
}

function isInstallationTableCompleted(record) {
  return getInstallationTableStatusValue(record) === "installed-completed";
}

function getInstallationTableStatusLabel(statusValue) {
  if (statusValue === "installed-completed") {
    return "Installed / Completed";
  }

  if (statusValue === "partial-installation") {
    return "Partial Installation";
  }

  return "Installation Blocker";
}

function getReadinessMeta(record) {
  const currentStageStatus = record?.saved2B
    ? getInstallationTableStatusLabel(
        getInstallationTableStatusValue(record),
      )
    : "Installation Pending";
  const blockerReason =
    currentStageStatus === "Installation Blocker"
      ? normalizeValue(record.blockerReason)
      : "";
  const billingStartDate = normalizeValue(record.billingStartDate);
  const validBillingStartDate =
    Boolean(billingStartDate) && isValidCalendarDateValue(billingStartDate);
  const billingMonthStartDate = validBillingStartDate
    ? getBillingMonthStartDate(billingStartDate)
    : null;
  const today = getTodayStart();
  const currentMonthStartDate = new Date(
    today.getFullYear(),
    today.getMonth(),
    1,
  );
  const normalizedCurrentStageStatus = normalizeValue(
    currentStageStatus,
  ).toLowerCase();
  const stage3Eligible =
    Boolean(billingMonthStartDate) &&
    normalizedCurrentStageStatus === "installed / completed" &&
    billingMonthStartDate <= currentMonthStartDate;
  const billingTreatment = validBillingStartDate
    ? getBillingTreatment(billingStartDate)
    : "";
  const billingReadinessStatus =
    getBillingReadinessStatus(billingStartDate, currentStageStatus) ||
    currentStageStatus;
  const billingMonthLabel = billingMonthStartDate
    ? new Intl.DateTimeFormat("en-US", {
        month: "long",
        year: "numeric",
      }).format(billingMonthStartDate)
    : "";

  return {
    currentStageStatus,
    blockerReason,
    validBillingStartDate,
    stage3Eligible,
    billingReadinessStatus,
    billingTreatment,
    billingMonthLabel,
  };
}

function getExpenseTotal(expenses = []) {
  return expenses.reduce((sum, expense) => sum + Number(expense.amount || 0), 0);
}

function getExpenseLedgerKey(record) {
  if (!record) {
    return "";
  }

  const billingCode = normalizeIdInput(record.billingCode);
  const screenCode = normalizeIdInput(record.screenCode);

  if (!screenCode) {
    return "";
  }

  // Expenses are saved against the individual Screen Code.
  // Complex totals are calculated by aggregating sibling ledgers.
  return `SITE::${billingCode}::${screenCode}`;
}

function getExpenseTimestamp(expense) {
  const value = normalizeValue(expense?.createdAt || expense?.updatedAt);
  const timestamp = value ? Date.parse(value) : NaN;
  return Number.isNaN(timestamp) ? null : timestamp;
}

function getExpenseCycleTimestamp(record) {
  const value = normalizeValue(record?.expenseCycleStartedAt);
  const timestamp = value ? Date.parse(value) : NaN;
  return Number.isNaN(timestamp) ? null : timestamp;
}

function getCurrentCycleExpenses(record) {
  const expenses = getOwnedExpenses(
    record?.installationExpenses,
    record?.screenCode,
    record?.screenName,
  );
  const cycleTimestamp = getExpenseCycleTimestamp(record);

  if (cycleTimestamp === null) return expenses;

  return expenses.filter((expense) => {
    const expenseTimestamp = getExpenseTimestamp(expense);
    return expenseTimestamp !== null && expenseTimestamp >= cycleTimestamp;
  });
}

function getHistoricalCycleExpenses(record) {
  const expenses = getOwnedExpenses(
    record?.installationExpenses,
    record?.screenCode,
    record?.screenName,
  );
  const cycleTimestamp = getExpenseCycleTimestamp(record);

  if (cycleTimestamp === null) return [];

  return expenses.filter((expense) => {
    const expenseTimestamp = getExpenseTimestamp(expense);
    return expenseTimestamp === null || expenseTimestamp < cycleTimestamp;
  });
}

function getRecordExpenseState(record) {
  const expenses = getCurrentCycleExpenses(record);
  const cycleTimestamp = getExpenseCycleTimestamp(record);
  const savedAtValue = normalizeValue(record?.installationExpensesSavedAt);
  const savedAtTimestamp = savedAtValue ? Date.parse(savedAtValue) : NaN;
  const currentCycleWasFinalized =
    cycleTimestamp === null ||
    (!Number.isNaN(savedAtTimestamp) && savedAtTimestamp >= cycleTimestamp);
  const applicable = currentCycleWasFinalized
    ? normalizeValue(record?.otherInstallationExpensesApplicable) ||
      (expenses.length > 0 ? "Yes" : "")
    : expenses.length > 0
      ? "Yes"
      : "";

  return { applicable, expenses };
}

function formatExpenseDateDisplay(value) {
  const parsed = parseDate(value);
  if (!parsed) {
    return value || "-";
  }

  const day = String(parsed.getUTCDate()).padStart(2, "0");
  const month = String(parsed.getUTCMonth() + 1).padStart(2, "0");
  const year = parsed.getUTCFullYear();
  return `${day}/${month}/${year}`;
}

function getInstallationSnapshot(record) {
  return {
    dateOfDispatch: normalizeValue(record.dateOfDispatch),
    installationDate: normalizeValue(record.installationDate),
    liveDate: normalizeValue(record.liveDate),
    trialPeriod: normalizeValue(record.trialPeriod),
    trialPeriodExtension: getNormalizedTrialPeriodExtension(record),
    blockerReason: normalizeValue(record.blockerReason),
    remarks: getNormalizedRemarks(record),
  };
}

function getInstallationValidationMessages(record) {
  const messages = [];
  const dateOfDispatch = normalizeValue(record?.dateOfDispatch);
  const installationDate = normalizeValue(record?.installationDate);
  const liveDate = normalizeValue(record?.liveDate);
  const billingStartDate = normalizeValue(record?.billingStartDate);

  const parsedDateOfDispatch = parseDate(dateOfDispatch);
  const parsedInstallationDate = parseDate(installationDate);
  const parsedLiveDate = parseDate(liveDate);
  const parsedBillingStartDate = parseDate(billingStartDate);

  if (
    dateOfDispatch &&
    installationDate &&
    parsedDateOfDispatch &&
    parsedInstallationDate &&
    parsedDateOfDispatch > parsedInstallationDate
  ) {
    messages.push("Date of Dispatch must be on or before Installation Date.");
  }

  if (dateOfDispatch && !isValidCalendarDateValue(dateOfDispatch)) {
    messages.push("Date of Dispatch: Please enter a valid calendar date.");
  }

  if (
    installationDate &&
    liveDate &&
    parsedInstallationDate &&
    parsedLiveDate &&
    parsedInstallationDate > parsedLiveDate
  ) {
    messages.push("Installation Date must be on or before Live Date.");
  }

  if (installationDate && isFutureDateValue(installationDate)) {
    messages.push("Installation Date must not be a future date.");
  }

  if (installationDate && !isValidCalendarDateValue(installationDate)) {
    messages.push("Installation Date: Please enter a valid calendar date.");
  }

  if (liveDate && isFutureDateValue(liveDate)) {
    messages.push("Live Date must not be a future date.");
  }

  if (liveDate && !isValidCalendarDateValue(liveDate)) {
    messages.push("Live Date: Please enter a valid calendar date.");
  }

  if (
    dateOfDispatch &&
    liveDate &&
    parsedDateOfDispatch &&
    parsedLiveDate &&
    parsedDateOfDispatch > parsedLiveDate
  ) {
    messages.push("Date of Dispatch must be on or before Live Date.");
  }

  if (
    liveDate &&
    billingStartDate &&
    parsedLiveDate &&
    parsedBillingStartDate &&
    parsedLiveDate > parsedBillingStartDate
  ) {
    messages.push("Live Date must be on or before Billing Start Date.");
  }

  return messages;
}

function haveInstallationDetailsChanged(previousSnapshot, nextSnapshot) {
  if (!previousSnapshot) {
    return false;
  }

  return Object.keys(nextSnapshot).some(
    (key) => previousSnapshot[key] !== nextSnapshot[key],
  );
}

function getExportColumnSet(stageFilter) {
  const detailedColumns = [
    "Billing Code / Customer Code",
    "Complex Code",
    "Screen Code",
    "Screen Name",
    "Location",
    "State",
    "Date of Dispatch",
    "Installation Date",
    "Live Date",
    "Trial Period",
    "Trial Period Extension",
    "Billing Start Date",
    "Current Stage Status",
    "Blocker Reason",
    "Remarks",
    "Installation Expenses Total",
  ];

  const shortColumns = [
    "Billing Code / Customer Code",
    "Complex Code",
    "Screen Code",
    "Screen Name",
    "Location",
    "State",
    "Date of Dispatch",
    "Current Stage Status",
    "Blocker Reason",
    "Remarks",
    "Installation Expenses Total",
  ];

  return ["Transport", "Installation Pending"].includes(stageFilter)
    ? shortColumns
    : detailedColumns;
}

function buildExportRow(record, expensesBySiteId, stageFilter) {
  const screenCode = normalizeValue(record.screenCode);
  const expenseState = screenCode
    ? expensesBySiteId[screenCode] || getDefaultExpenseState()
    : getDefaultExpenseState();
  const expenseTotal = getExpenseTotal(expenseState.expenses);
  const currentStageStatus = resolveStage2Status(record);
  const detailed = stageFilter === "All" || !["Transport", "Installation Pending"].includes(stageFilter);

  const row = {
    "Billing Code / Customer Code": normalizeValue(record.billingCode),
    "Complex Code": normalizeValue(record.complexCode),
    "Screen Code": normalizeValue(record.screenCode),
    "Screen Name": normalizeValue(record.screenName),
    Location: normalizeValue(record.location),
    State: normalizeValue(record.state),
    "Date of Dispatch": normalizeValue(record.dateOfDispatch),
    "Current Stage Status": currentStageStatus,
    "Blocker Reason": normalizeValue(record.blockerReason),
    Remarks: getNormalizedRemarks(record),
    "Installation Expenses Total": expenseTotal ? `ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¹${formatExpenseAmount(expenseTotal)}` : "ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¹0.00",
  };

  if (detailed) {
    row["Installation Date"] = normalizeValue(record.installationDate);
    row["Live Date"] = normalizeValue(record.liveDate);
    row["Trial Period"] = normalizeValue(record.trialPeriod);
    row["Trial Period Extension"] = getNormalizedTrialPeriodExtension(record);
    row["Billing Start Date"] = normalizeValue(record.billingStartDate);
  }

  return row;
}

function getTodayStart() {
  const today = new Date();
  return new Date(today.getFullYear(), today.getMonth(), today.getDate());
}

function Installations({
  stage2Records = [],
  setStage2Records = () => {},
  referenceStage3Records = [],
  onMoveToStage3 = () => false,
  isMinimized = false,
  onReturnToStage1 = () => {},
  focusRecordId = "",
  onFocusRecordConsumed = () => {},
}) {
  const [searchTerm, setSearchTerm] = useState("");
  const [readinessSearchTerm, setReadinessSearchTerm] = useState("");
  const [readinessSortOrder, setReadinessSortOrder] = useState("desc");
  const [readinessSortKey, setReadinessSortKey] = useState("");
  const [selectedReadyRecordIds, setSelectedReadyRecordIds] = useState([]);
  const [readinessStageFilter, setReadinessStageFilter] = useState("all");
  const [readinessFromDate, setReadinessFromDate] = useState("");
  const [readinessToDate, setReadinessToDate] = useState("");
  const [activeRecordId, setActiveRecordId] = useState("");
  const [expenseType, setExpenseType] = useState("");
  const [expenseDate, setExpenseDate] = useState("");
  const [expenseAmount, setExpenseAmount] = useState("");
  const [expenseComments, setExpenseComments] = useState("");
  const [expenseMessage, setExpenseMessage] = useState("");
  const [editingExpenseId, setEditingExpenseId] = useState("");
  const [expensesBySiteId, setExpensesBySiteId] = useState({});
  const [persistedExpenseReferenceRecords, setPersistedExpenseReferenceRecords] =
    useState([]);
  const [downloadStage, setDownloadStage] = useState("All");
  const [downloadFromDate, setDownloadFromDate] = useState("");
  const [downloadToDate, setDownloadToDate] = useState("");
  const [readinessDetailsRecord, setReadinessDetailsRecord] = useState(null);
  const [activeInstallationRecordId, setActiveInstallationRecordId] = useState("");
  const [installationSiteSearchTerm, setInstallationSiteSearchTerm] = useState("");
  const [lastSaved2BRecordId, setLastSaved2BRecordId] = useState("");
  const [copySourceRecordId, setCopySourceRecordId] = useState("");
  const [copyTargetRecordId, setCopyTargetRecordId] = useState("");
  
  const readinessSelectAllRef = useRef(null);
  const siteSelectionRef = useRef(null);
  const complexSitesRef = useRef(null);
  const billingReadinessQueueRef = useRef(null);
  const dispatchDateRef = useRef(null);
  const installationDateRef = useRef(null);
  const liveDateRef = useRef(null);

  const normalizedSearch = searchTerm.trim().toLowerCase();

  useEffect(() => {
    let cancelled = false;

    async function restoreExpenseReferences() {
      try {
        const response = await fetch(`${BILLING_API_BASE}/sites`);
        const body = await response.json().catch(() => ({}));
        if (!response.ok) {
          throw new Error(body?.message || body?.error || "Unable to restore expense references.");
        }

        const references = (Array.isArray(body?.data) ? body.data : [])
          .map((site) => {
            const stage1 = site?.stage1Data && typeof site.stage1Data === "object"
              ? site.stage1Data
              : {};
            const stage2 = site?.stage2Data && typeof site.stage2Data === "object"
              ? site.stage2Data
              : {};
            const stage3 = site?.stage3Data && typeof site.stage3Data === "object"
              ? site.stage3Data
              : {};
            const screenCode = normalizeIdInput(site?.siteId);
            const screenName = normalizeValue(site?.screenName);
            const snapshot = { ...stage1, ...stage2, ...stage3 };

            if (!screenCode) return null;

            return {
              ...snapshot,
              backendSiteId: site.id,
              recordId: snapshot.recordId || site.id,
              billingCode: site.billingId || "",
              complexCode: site.complexId || "",
              screenCode,
              screenName,
              installationExpenses: getOwnedExpenses(
                stage3.installationExpenses ?? stage2.installationExpenses,
                screenCode,
                screenName,
              ),
              workflowReferenceOnly: true,
              workflowReferenceStage: site.processingStatus || "Persisted",
            };
          })
          .filter(Boolean);

        if (!cancelled) setPersistedExpenseReferenceRecords(references);
      } catch (error) {
        if (!cancelled) {
          console.error("Unable to restore persisted expense references:", error);
        }
      }
    }

    restoreExpenseReferences();
    return () => {
      cancelled = true;
    };
  }, []);

  const filteredRecords = useMemo(
    () =>
      stage2Records.filter((record) => {
        const recordText = [
          record.billingCode,
          record.complexCode,
          record.screenCode,
          record.screenName,
          record.location,
        ]
          .join(" ")
          .toLowerCase();

        return normalizedSearch === "" || recordText.includes(normalizedSearch);
      }),
    [stage2Records, normalizedSearch],
  );

  const installationContextRecords = useMemo(() => {
    const bySiteId = new Map();

    // Include every persisted sibling as a read-only reference so a screen
    // returned to Stage 1 cannot make its saved expense disappear from the
    // current complex ledger.
    persistedExpenseReferenceRecords.forEach((record) => {
      const screenCode = normalizeIdInput(record?.screenCode);
      if (screenCode) bySiteId.set(screenCode, record);
    });

    // Stage 3 records are reference-only in the Installation workspace.
    referenceStage3Records.forEach((record) => {
      const screenCode = normalizeIdInput(record?.screenCode);

      if (!screenCode) {
        return;
      }

      bySiteId.set(screenCode, {
        ...record,
        workflowReferenceOnly: true,
        workflowReferenceStage: "Stage 3",
      });
    });

    // A live Stage 2 record always wins over reference history.
    stage2Records.forEach((record) => {
      const screenCode = normalizeIdInput(record?.screenCode);

      if (!screenCode) {
        return;
      }

      bySiteId.set(screenCode, {
        ...record,
        workflowReferenceOnly: false,
        workflowReferenceStage: "Stage 2",
      });
    });

    return Array.from(bySiteId.values());
  }, [persistedExpenseReferenceRecords, referenceStage3Records, stage2Records]);

  const activeRecord = useMemo(
    () =>
      stage2Records.find((record) => record.recordId === activeRecordId) || null,
    [activeRecordId, stage2Records],
  );

  const activeInstallationRecord = useMemo(
    () =>
      installationContextRecords.find(
        (record) => record.recordId === activeInstallationRecordId,
      ) || null,
    [activeInstallationRecordId, installationContextRecords],
  );

  const normalizedInstallationSiteSearch =
    installationSiteSearchTerm.trim().toLowerCase();

  const installationGroupRecords = useMemo(() => {
    if (!activeInstallationRecord) {
      return [];
    }

    const selectedBillingId = normalizeIdInput(
      activeInstallationRecord.billingCode,
    );
    const selectedComplexId = normalizeIdInput(
      activeInstallationRecord.complexCode,
      { allowStandaloneBlank: true },
    );

    if (selectedComplexId) {
      return installationContextRecords.filter((record) => {
        const recordBillingId = normalizeIdInput(record.billingCode);
        const recordComplexId = normalizeIdInput(record.complexCode, {
          allowStandaloneBlank: true,
        });

        return (
          recordBillingId === selectedBillingId &&
          recordComplexId === selectedComplexId
        );
      });
    }

    return installationContextRecords.filter(
      (record) => record.recordId === activeInstallationRecord.recordId,
    );
  }, [activeInstallationRecord, installationContextRecords]);

  const installationVisibleRecords = useMemo(() => {
    if (!activeInstallationRecord) {
      return [];
    }

    if (!normalizedInstallationSiteSearch) {
      return installationGroupRecords;
    }

    return installationGroupRecords.filter((record) => {
      const searchableText = [
        record.billingCode,
        record.complexCode,
        record.billingName,
        record.screenCode,
        record.screenName,
        record.location,
        record.state,
      ]
        .map((value) => normalizeValue(value).toLowerCase())
        .join(" ");

      return searchableText.includes(normalizedInstallationSiteSearch);
    });
  }, [
    activeInstallationRecord,
    installationGroupRecords,
    normalizedInstallationSiteSearch,
  ]);

  const activeInstallationStatusValue = activeInstallationRecord
    ? getInstallationTableStatusValue(activeInstallationRecord)
    : "";
  const activeInstallationStatusCompleted = activeInstallationRecord
    ? isInstallationTableCompleted(activeInstallationRecord)
    : false;

  const completedInstallationCount = useMemo(
    () =>
      installationGroupRecords.filter(
        (record) =>
          record.saved2B &&
          getInstallationTableStatusValue(record) === "installed-completed",
      ).length,
    [installationGroupRecords],
  );

  const savedInstallationCount = useMemo(
    () => installationGroupRecords.filter((record) => record.saved2B).length,
    [installationGroupRecords],
  );

  const expenseLedgerKey = getExpenseLedgerKey(activeInstallationRecord);
  const persistedActiveExpenseState = getRecordExpenseState(
    activeInstallationRecord,
  );
  const activeSiteExpenseState = expenseLedgerKey
    ? expensesBySiteId[expenseLedgerKey] || persistedActiveExpenseState
    : getDefaultExpenseState();
  const activeSiteExpenseApplicable = activeSiteExpenseState.applicable || "";
  const activeSiteExpenses = getOwnedExpenses(
    activeSiteExpenseState.expenses,
    activeInstallationRecord?.screenCode,
    activeInstallationRecord?.screenName,
  );
  const isActiveInstallationReadOnly =
    activeInstallationRecord?.workflowReferenceOnly === true;

  const previousComplexExpenses = useMemo(() => {
    if (!activeInstallationRecord) {
      return [];
    }

    const activeScreenCode = normalizeIdInput(
      activeInstallationRecord.screenCode,
    );
    const byExpenseId = new Map();

    getHistoricalCycleExpenses(activeInstallationRecord).forEach((expense, index) => {
      const expenseId =
        normalizeValue(expense?.expenseId) ||
        `${activeScreenCode}-history-${index}-${expense?.expenseDate || ""}-${expense?.amount || ""}`;

      byExpenseId.set(expenseId, {
        ...expense,
        expenseId,
        screenCode:
          normalizeValue(expense?.screenCode) ||
          normalizeValue(activeInstallationRecord?.screenCode),
        screenName:
          normalizeValue(expense?.screenName) ||
          normalizeValue(activeInstallationRecord?.screenName),
        readOnly: true,
      });
    });

    installationGroupRecords.forEach((record) => {
      const screenCode = normalizeIdInput(record?.screenCode);

      if (!screenCode || screenCode === activeScreenCode) {
        return;
      }

      const recordExpenses = getOwnedExpenses(
        record?.installationExpenses,
        record?.screenCode,
        record?.screenName,
      );

      recordExpenses.forEach((expense, index) => {
        const expenseId =
          normalizeValue(expense?.expenseId) ||
          `${screenCode}-legacy-${index}-${expense?.expenseDate || ""}-${expense?.amount || ""}`;

        if (!byExpenseId.has(expenseId)) {
          byExpenseId.set(expenseId, {
            ...expense,
            expenseId,
            screenCode:
              normalizeValue(expense?.screenCode) ||
              normalizeValue(record?.screenCode),
            screenName:
              normalizeValue(expense?.screenName) ||
              normalizeValue(record?.screenName),
            readOnly: true,
          });
        }
      });
    });

    return Array.from(byExpenseId.values());
  }, [activeInstallationRecord, installationGroupRecords]);

  const previousExpensesTotal = getExpenseTotal(previousComplexExpenses);
  const currentSiteExpensesTotal = getExpenseTotal(activeSiteExpenses);
  const combinedExpensesTotal =
    previousExpensesTotal + currentSiteExpensesTotal;

  const hasAnySavedInstallationDetails = useMemo(
    () =>
      installationGroupRecords.some(
        (record) =>
          record.saved2B && isInstallationTableCompleted(record),
      ),
    [installationGroupRecords],
  );

  const canEditExpenses =
    hasAnySavedInstallationDetails &&
    !isActiveInstallationReadOnly;

  const activeCurrentStageStatus = activeRecord
    ? activeRecord?.saved2B
      ? getInstallationTableStatusLabel(
          getInstallationTableStatusValue(activeRecord),
        )
      : "Installation Pending"
    : "";
  const installationValidationMessages = useMemo(
    () => (activeRecord ? getInstallationValidationMessages(activeRecord) : []),
    [activeRecord],
  );

  const downloadFilteredRecords = useMemo(
    () =>
      stage2Records.filter((record) => {
        const currentStageStatus = resolveStage2Status(record);

        if (downloadStage !== "All" && currentStageStatus !== downloadStage) {
          return false;
        }

        const dispatchDateValue = normalizeValue(record.dateOfDispatch);
        if (!dispatchDateValue) {
          return downloadFromDate === "" && downloadToDate === "";
        }

        const parsedDispatchDate = parseDate(dispatchDateValue);
        if (!parsedDispatchDate) {
          return false;
        }

        if (downloadFromDate) {
          const parsedFromDate = parseDate(downloadFromDate);
          if (parsedFromDate && parsedDispatchDate < parsedFromDate) {
            return false;
          }
        }

        if (downloadToDate) {
          const parsedToDate = parseDate(downloadToDate);
          if (parsedToDate && parsedDispatchDate > parsedToDate) {
            return false;
          }
        }

        return true;
      }),
    [downloadFromDate, downloadStage, downloadToDate, stage2Records],
  );

  const downloadHasMatches = downloadFilteredRecords.length > 0;
  const downloadColumnHeaders = getExportColumnSet(downloadStage);

  const readinessRows = useMemo(
    () =>
      stage2Records.map((record) => ({
        ...record,
        ...getReadinessMeta(record),
      })),
    [stage2Records],
  );

  const filteredReadinessRows = useMemo(() => {
    const normalizedReadinessSearch = readinessSearchTerm.trim().toLowerCase();
    const parsedReadinessFromDate = readinessFromDate
      ? parseDate(readinessFromDate)
      : null;
    const parsedReadinessToDate = readinessToDate ? parseDate(readinessToDate) : null;

    const filtered = readinessRows.filter((record) => {
      const readinessStageKey = getReadinessStageKey(record);
      if (
        readinessStageFilter !== "all" &&
        readinessStageKey !== readinessStageFilter
      ) {
        return false;
      }

      if (normalizedReadinessSearch) {
        const searchableText = [
          record.billingCode,
          record.complexCode,
          record.screenCode,
          record.screenName,
          record.location,
          record.state,
          record.billingReadinessStatus,
          record.currentStageStatus,
        ]
          .join(" ")
          .toLowerCase();

        if (!searchableText.includes(normalizedReadinessSearch)) {
          return false;
        }
      }

      const billingStartDate = normalizeValue(record.billingStartDate);
      if (parsedReadinessFromDate || parsedReadinessToDate) {
        if (!billingStartDate) {
          return false;
        }

        const parsedBillingStartDate = parseDate(billingStartDate);
        if (!parsedBillingStartDate) {
          return false;
        }

        if (
          parsedReadinessFromDate &&
          parsedBillingStartDate < parsedReadinessFromDate
        ) {
          return false;
        }

        if (
          parsedReadinessToDate &&
          parsedBillingStartDate > parsedReadinessToDate
        ) {
          return false;
        }
      }

      return true;
    });

    return filtered.sort((left, right) => {
      const valueFor = (record) => {
        if (!readinessSortKey) return normalizeValue(record.billingStartDate);
        if (readinessSortKey === "billingReadiness") return formatReadinessBadgeLabel(record);
        return record?.[readinessSortKey];
      };
      const comparison = compareColumnValues(valueFor(left), valueFor(right));
      return readinessSortOrder === "asc" ? comparison : -comparison;
    });
  }, [
    readinessFromDate,
    readinessRows,
    readinessSearchTerm,
    readinessStageFilter,
    readinessToDate,
    readinessSortOrder,
    readinessSortKey,
  ]);

  function handleReadinessColumnSort(key) {
    if (readinessSortKey === key) setReadinessSortOrder((current) => (current === "asc" ? "desc" : "asc"));
    else { setReadinessSortKey(key); setReadinessSortOrder("asc"); }
  }

  const eligibleReadinessRows = useMemo(
    () => filteredReadinessRows.filter((record) => record.stage3Eligible),
    [filteredReadinessRows],
  );

  const eligibleReadinessRowIds = useMemo(
    () => new Set(eligibleReadinessRows.map((record) => record.recordId)),
    [eligibleReadinessRows],
  );

  const hasSelectedReadyRecords = selectedReadyRecordIds.length > 0;
  const hasVisibleReadyRecords = eligibleReadinessRows.length > 0;
  const areAllVisibleReadyRecordsSelected =
    hasVisibleReadyRecords &&
    eligibleReadinessRows.every((record) =>
      selectedReadyRecordIds.includes(record.recordId),
    );
  const areSomeVisibleReadyRecordsSelected =
    hasVisibleReadyRecords &&
    eligibleReadinessRows.some((record) =>
      selectedReadyRecordIds.includes(record.recordId),
    ) &&
    !areAllVisibleReadyRecordsSelected;

  useEffect(() => {
    if (!readinessSelectAllRef.current) {
      return;
    }

    readinessSelectAllRef.current.indeterminate =
      areSomeVisibleReadyRecordsSelected;
  }, [areSomeVisibleReadyRecordsSelected]);

  useEffect(() => {
    setSelectedReadyRecordIds((currentIds) =>
      currentIds.filter((recordId) => eligibleReadinessRowIds.has(recordId)),
    );
  }, [eligibleReadinessRowIds]);

  useEffect(() => {
    setSelectedReadyRecordIds([]);
  }, [readinessStageFilter]);

  useEffect(() => {
    if (focusRecordId) {
      const focusRecord = stage2Records.find(
        (record) => record.recordId === focusRecordId,
      );

      if (focusRecord) {
        handleEditRecord(focusRecordId);
        onFocusRecordConsumed();
        return;
      }
    }

    window.requestAnimationFrame(() => {
      guidedScrollToElement(billingReadinessQueueRef.current);
    });
  }, [focusRecordId]);

  useEffect(() => {
    if (
      activeRecordId &&
      !stage2Records.some((record) => record.recordId === activeRecordId)
    ) {
      setActiveRecordId("");
    }
  }, [activeRecordId, stage2Records]);

  useEffect(() => {
    if (
      activeInstallationRecordId &&
      !installationContextRecords.some(
        (record) => record.recordId === activeInstallationRecordId,
      )
    ) {
      setActiveInstallationRecordId("");
    }
  }, [activeInstallationRecordId, installationContextRecords]);

  useEffect(() => {
    if (
      copySourceRecordId &&
      !installationContextRecords.some(
        (record) => record.recordId === copySourceRecordId,
      )
    ) {
      setCopySourceRecordId("");
    }

    if (
      copyTargetRecordId &&
      !stage2Records.some((record) => record.recordId === copyTargetRecordId)
    ) {
      setCopyTargetRecordId("");
    }
  }, [
    copySourceRecordId,
    copyTargetRecordId,
    installationContextRecords,
    stage2Records,
  ]);

  useEffect(() => {
    setExpenseDate("");
    setExpenseAmount("");
    setExpenseComments("");
    setExpenseMessage("");
    setEditingExpenseId("");
  }, [expenseLedgerKey]);

  useEffect(() => {
    if (!activeRecord || isBillingCommercialLocked(activeRecord)) {
      return;
    }

    const nextBillingStartDate = calculateBillingStartDate(
      activeRecord.liveDate,
      activeRecord.trialPeriod,
      getNormalizedTrialPeriodExtension(activeRecord),
    );

    if ((activeRecord.billingStartDate || "") === nextBillingStartDate) {
      return;
    }

    setStage2Records((currentRecords) =>
      currentRecords.map((record) =>
        record.recordId === activeRecord.recordId
          ? { ...record, billingStartDate: nextBillingStartDate }
          : record,
      ),
    );
  }, [
    activeRecord?.billingStartDate,
    activeRecord?.liveDate,
    activeRecord?.recordId,
    activeRecord?.trialExtension,
    activeRecord?.trialPeriodExtension,
    activeRecord?.totalTrialPeriodExtension,
    activeRecord?.trialPeriod,
    activeRecord,
    setStage2Records,
  ]);

  function updateActiveRecord(field, value) {
    if (!activeRecord) {
      return;
    }

    if (
      isBillingCommercialLocked(activeRecord) &&
      billingLockedInstallationFields.has(field)
    ) {
      alert(
        "This billing-cycle field is locked because the data was already shared with the Billing Team.",
      );
      return;
    }

    setStage2Records((currentRecords) =>
      currentRecords.map((record) =>
        record.recordId === activeRecord.recordId
          ? {
              ...record,
              [field]: value,
              ...(field === "trialPeriodExtension" ||
              field === "totalTrialPeriodExtension" ||
              field === "trialExtension"
                ? {
                    trialPeriodExtension: value,
                    totalTrialPeriodExtension: value,
                    trialExtension: value,
                  }
                : {}),
              ...(field === "remarks" || field === "extensionRemarks"
                ? {
                    remarks: value,
                    extensionRemarks: value,
                  }
                : {}),
              ...((field === "liveDate" ||
                field === "trialPeriod" ||
                field === "trialPeriodExtension" ||
                field === "totalTrialPeriodExtension" ||
                field === "trialExtension")
                ? {
                    billingStartDate: calculateBillingStartDate(
                      field === "liveDate" ? value : record.liveDate,
                      field === "trialPeriod" ? value : record.trialPeriod,
                      field === "trialPeriodExtension" ||
                      field === "totalTrialPeriodExtension" ||
                      field === "trialExtension"
                        ? value
                        : getNormalizedTrialPeriodExtension(record),
                    ),
                  }
                : {}),
            }
          : record,
      ),
    );
  }

  function updateInstallationTableRecord(recordId, field, value) {
    setStage2Records((currentRecords) =>
      currentRecords.map((record) => {
        if (record.recordId !== recordId) {
          return record;
        }

        if (
          isBillingCommercialLocked(record) &&
          billingLockedInstallationFields.has(field)
        ) {
          return record;
        }

        const nextRecord = { ...record };

        if (
          field === "trialPeriodExtension" ||
          field === "totalTrialPeriodExtension" ||
          field === "trialExtension"
        ) {
          nextRecord.trialPeriodExtension = value;
          nextRecord.totalTrialPeriodExtension = value;
          nextRecord.trialExtension = value;
        } else {
          nextRecord[field] = value;
        }

        if (field === "installationStageStatus") {
          nextRecord.installationDetailsCompleted =
            value === "installed-completed";

          if (value !== "installation-blocker") {
            nextRecord.installationBlocker = "";
            nextRecord.blockerReason = "";
          }
        }

        if (field === "installationBlocker") {
          nextRecord.installationBlocker = value;
        }

        if (field === "blockerReason") {
          nextRecord.blockerReason = value;
        }

        if (
          field === "installationDate" ||
          field === "liveDate" ||
          field === "billingStartDate" ||
          field === "trialPeriod" ||
          field === "trialPeriodExtension" ||
          field === "totalTrialPeriodExtension" ||
          field === "trialExtension"
        ) {
          nextRecord.billingStartDate = calculateInstallationBillingDate(
            nextRecord,
          );
        }

        return nextRecord;
      }),
    );
  }

  function handleInstallationTableStatusChange(recordId, value) {
    updateInstallationTableRecord(recordId, "installationStageStatus", value);
  }

  function handleSelectCopySource(recordId) {
    if (installationGroupRecords.length <= 1) {
      return;
    }

    const sourceRecord = installationGroupRecords.find(
      (record) => record.recordId === recordId,
    );

    if (
      !sourceRecord ||
      !sourceRecord.saved2B ||
      !isInstallationTableCompleted(sourceRecord)
    ) {
      return;
    }

    if (copySourceRecordId === recordId) {
      setCopySourceRecordId("");
      setCopyTargetRecordId("");
      return;
    }

    setCopySourceRecordId(recordId);
    setCopyTargetRecordId("");
  }

  function handleInstallationSiteCardClick(record) {
    if (!record) {
      return;
    }

    if (!copySourceRecordId) {
      const canUseAsCopySource =
        installationGroupRecords.length > 1 &&
        record.saved2B &&
        isInstallationTableCompleted(record);

      if (record.workflowReferenceOnly) {
        setActiveRecordId("");
      } else {
        setActiveRecordId(record.recordId);
      }

      setActiveInstallationRecordId(record.recordId);
      setInstallationSiteSearchTerm("");

      // Previous Stage 2 behaviour: clicking a completed Complex-site card
      // selects that card itself as Copy From. No dropdown is required.
      if (canUseAsCopySource) {
        handleSelectCopySource(record.recordId);
      }
      return;
    }

    if (record.recordId === copySourceRecordId) {
      // Clicking the source card again cancels Copy From mode.
      handleSelectCopySource(record.recordId);
      setActiveRecordId(record.workflowReferenceOnly ? "" : record.recordId);
      setActiveInstallationRecordId(record.recordId);
      setInstallationSiteSearchTerm("");
      return;
    }

    if (isBillingCommercialLocked(record)) {
      alert(
        "Installation/Billing dates cannot be copied into this site because its billing cycle is locked after Billing Team submission.",
      );
      return;
    }

    const sourceRecord = installationGroupRecords.find(
      (candidate) => candidate.recordId === copySourceRecordId,
    );

    if (!sourceRecord) {
      setCopySourceRecordId("");
      setCopyTargetRecordId("");
      return;
    }

    const sourceComplexCode = normalizeIdInput(sourceRecord.complexCode, {
      allowStandaloneBlank: true,
    });
    const targetComplexCode = normalizeIdInput(record.complexCode, {
      allowStandaloneBlank: true,
    });
    const sourceBillingCode = normalizeIdInput(sourceRecord.billingCode);
    const targetBillingCode = normalizeIdInput(record.billingCode);

    if (
      !sourceComplexCode ||
      sourceComplexCode !== targetComplexCode ||
      sourceBillingCode !== targetBillingCode
    ) {
      alert("Installation details can be copied only within the same Complex.");
      return;
    }

    const hasExistingTargetValues = [
      record.dateOfDispatch,
      record.installationDate,
      record.liveDate,
      record.trialPeriod,
      getNormalizedTrialPeriodExtension(record),
    ].some((value) => Boolean(normalizeValue(value)));

    const confirmed = window.confirm(
      `Copy Dispatch Date, Installation Date, Live Date, Trial Period and Additional TP from ${
        sourceRecord.screenCode || "the source site"
      } to ${record.screenCode || "the destination site"}?${
        hasExistingTargetValues
          ? "\n\nExisting values in these five destination fields will be replaced."
          : ""
      }`,
    );

    if (!confirmed) {
      return;
    }

    const copiedTrialPeriodExtension =
      getNormalizedTrialPeriodExtension(sourceRecord);

    setStage2Records((currentRecords) =>
      currentRecords.map((candidate) => {
        if (candidate.recordId !== record.recordId) {
          return candidate;
        }

        const nextRecord = {
          ...candidate,
          dateOfDispatch: normalizeValue(sourceRecord.dateOfDispatch),
          installationDate: normalizeValue(sourceRecord.installationDate),
          liveDate: normalizeValue(sourceRecord.liveDate),
          trialPeriod: normalizeValue(sourceRecord.trialPeriod),
          trialPeriodExtension: copiedTrialPeriodExtension,
          totalTrialPeriodExtension: copiedTrialPeriodExtension,
          trialExtension: copiedTrialPeriodExtension,
        };

        nextRecord.billingStartDate = calculateInstallationBillingDate(
          nextRecord,
        );

        return nextRecord;
      }),
    );

    setCopyTargetRecordId(record.recordId);
    setActiveRecordId(record.recordId);
    setActiveInstallationRecordId(record.recordId);
    setInstallationSiteSearchTerm("");
  }

  async function handleSaveInstallationTableChanges() {
    if (stage2Records.length === 0) {
      alert("No installation records available.");
      return;
    }

    if (!activeInstallationRecord) {
      alert("Select a site before saving installation details.");
      return;
    }

    if (activeInstallationRecord.workflowReferenceOnly) {
      alert("Earlier installed/downstream site information is read-only.");
      return;
    }

    const recordId = activeInstallationRecord.recordId;
    const dateOfDispatch = normalizeValue(activeInstallationRecord.dateOfDispatch);
    const installationDate = normalizeValue(activeInstallationRecord.installationDate);
    const liveDate = normalizeValue(activeInstallationRecord.liveDate);
    const billingStartDate = normalizeValue(activeInstallationRecord.billingStartDate);
    const currentStatus = getInstallationTableStatusValue(activeInstallationRecord);

    if (dateOfDispatch && !isValidCalendarDateValue(dateOfDispatch)) {
      alert("Dispatch Date: Please enter a valid calendar date.");
      return;
    }

    if (dateOfDispatch && isFutureDateValue(dateOfDispatch)) {
      alert("Dispatch Date must not be a future date.");
      return;
    }

    if (
      installationDate &&
      !isValidCalendarDateValue(installationDate)
    ) {
      alert("Installation Date: Please enter a valid calendar date.");
      return;
    }

    if (
      liveDate &&
      !isValidCalendarDateValue(liveDate)
    ) {
      alert("Live Date: Please enter a valid calendar date.");
      return;
    }

    if (
      installationDate &&
      isFutureDateValue(installationDate)
    ) {
      alert("Installation Date must not be a future date.");
      return;
    }

    if (
      liveDate &&
      isFutureDateValue(liveDate)
    ) {
      alert("Live Date must not be a future date.");
      return;
    }

    if (
      dateOfDispatch &&
      installationDate &&
      parseDate(dateOfDispatch) > parseDate(installationDate)
    ) {
      alert("Dispatch Date must be on or before Installation Date.");
      return;
    }

    if (
      installationDate &&
      liveDate &&
      parseDate(installationDate) > parseDate(liveDate)
    ) {
      alert("Installation Date must be on or before Live Date.");
      return;
    }

    if (
      dateOfDispatch &&
      liveDate &&
      parseDate(dateOfDispatch) > parseDate(liveDate)
    ) {
      alert("Dispatch Date must be on or before Live Date.");
      return;
    }

    let nextStatus;

    if (installationDate && liveDate && billingStartDate) {
      nextStatus = "installed-completed";
    } else if (installationDate) {
      nextStatus = "partial-installation";
    } else if (currentStatus === "partial-installation") {
      nextStatus = "partial-installation";
    } else {
      nextStatus = "installation-blocker";
    }

    const savedStatusLabel =
      nextStatus === "installed-completed"
        ? "Installed / Completed"
        : nextStatus === "partial-installation"
          ? "Partial Installation"
          : "Installation Blocker";

    if (
      nextStatus === "installation-blocker" &&
      !normalizeValue(activeInstallationRecord.blockerReason)
    ) {
      alert("Blocker Reason is required for Installation Blocker.");
      return;
    }

    const savedRecord = {
      ...activeInstallationRecord,
      installationStageStatus: nextStatus,
      installationDetailsCompleted: nextStatus === "installed-completed",
      saved2B: true,
      saved2BAt: new Date().toISOString(),
    };

    if (nextStatus !== "installation-blocker") {
      savedRecord.installationBlocker = "";
      savedRecord.blockerReason = "";
    }

    let persisted;
    try {
      persisted = await persistStage2Record(savedRecord);
    } catch (error) {
      alert(error?.message || "Stage 2 Installation Details could not be saved.");
      return;
    }

    setStage2Records((currentRecords) =>
      currentRecords.map((record) =>
        record.recordId === recordId
          ? { ...savedRecord, backendSiteId: persisted.backendSiteId }
          : record,
      ),
    );

    setLastSaved2BRecordId(recordId);
    setCopySourceRecordId("");
    setCopyTargetRecordId("");
    setExpenseMessage(
      `Installation Details saved with status: ${savedStatusLabel}. Other Installation Expenses are now available.`,
    );
  }

  function handleValidatedDateChange(field) {
    return (event) => {
      const nextValue = event.target.value;

      if (!nextValue) {
        updateActiveRecord(field, "");
        return;
      }

      updateActiveRecord(field, nextValue);
    };
  }

  function handleReturnToStage1() {
    if (!activeRecord) {
      return;
    }

    onReturnToStage1(activeRecord);
  }

  function handleSelectSearchResult(recordId) {
    setActiveRecordId(recordId);
  }

  function updateActiveSiteExpenseState(updater) {
    if (!expenseLedgerKey) {
      return;
    }

    setExpensesBySiteId((currentExpenses) => {
      const currentState =
        currentExpenses[expenseLedgerKey] || getDefaultExpenseState();
      const nextState =
        typeof updater === "function" ? updater(currentState) : updater;

      return {
        ...currentExpenses,
        [expenseLedgerKey]: {
          ...currentState,
          ...nextState,
        },
      };
    });
  }

  function handleExpenseApplicabilityChange(event) {
    if (!canEditExpenses || !expenseLedgerKey) {
      return;
    }

    const nextApplicable = event.target.value;

    if (nextApplicable === "No" && activeSiteExpenses.length > 0) {
      setExpenseMessage(
        "Saved expense history exists. Existing financial entries cannot be removed by changing this to No.",
      );
      return;
    }

    updateActiveSiteExpenseState({ applicable: nextApplicable });
    setExpenseMessage("");
  }

  function resetExpenseEditor() {
    setExpenseType("");
    setExpenseDate("");
    setExpenseAmount("");
    setExpenseComments("");
    setEditingExpenseId("");
  }

  function handleSaveExpense() {
    if (!canEditExpenses || !expenseLedgerKey) {
      setExpenseMessage(
        "Save at least one screen's Installation Details before adding expenses.",
      );
      return;
    }

    if (activeSiteExpenseApplicable !== "Yes") {
      setExpenseMessage("Select Yes for Other Installation Expenses first.");
      return;
    }

    const nextType = normalizeValue(expenseType);
    const nextDate = normalizeValue(expenseDate);
    const nextAmount = normalizeValue(expenseAmount);
    const nextComments = normalizeValue(expenseComments);

    if (!nextType) {
      setExpenseMessage("Expense Type is mandatory.");
      return;
    }

    if (!nextDate) {
      setExpenseMessage("Expense Date is mandatory.");
      return;
    }

    if (!isValidCalendarDateValue(nextDate)) {
      setExpenseMessage("Expense Date: Please enter a valid calendar date.");
      return;
    }

    if (isFutureDateValue(nextDate)) {
      setExpenseMessage("Expense Date must not be a future date.");
      return;
    }

    if (!isValidExpenseAmount(nextAmount)) {
      setExpenseMessage("Amount must be numeric and greater than 0.");
      return;
    }

    if (!nextComments) {
      setExpenseMessage("Comments are mandatory for every expense entry.");
      return;
    }

    const now = new Date().toISOString();

    updateActiveSiteExpenseState((currentState) => {
      const currentExpenses = dedupeExpenses(
        currentState.expenses,
        activeInstallationRecord?.screenCode,
      );

      if (editingExpenseId) {
        return {
          applicable: "Yes",
          expenses: currentExpenses.map((expense) =>
            expense.expenseId === editingExpenseId
              ? {
                  ...expense,
                  expenseType: nextType,
                  expenseDate: nextDate,
                  amount: Number(nextAmount),
                  comments: nextComments,
                  updatedAt: now,
                }
              : expense,
          ),
        };
      }

      return {
        applicable: "Yes",
        expenses: [
          ...currentExpenses,
          {
            expenseId: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
            screenCode: normalizeValue(activeInstallationRecord?.screenCode),
            screenName: normalizeValue(activeInstallationRecord?.screenName),
            expenseType: nextType,
            expenseDate: nextDate,
            amount: Number(nextAmount),
            comments: nextComments,
            createdAt: now,
            updatedAt: now,
          },
        ],
      };
    });

    setExpenseMessage(
      editingExpenseId ? "Expense entry updated." : "Expense entry added.",
    );
    resetExpenseEditor();
  }

  function handleEditExpense(expenseId) {
    const expense = activeSiteExpenses.find(
      (entry) => entry.expenseId === expenseId,
    );

    if (!expense) {
      return;
    }

    setEditingExpenseId(expenseId);
    setExpenseType(expense.expenseType || "");
    setExpenseDate(expense.expenseDate || "");
    setExpenseAmount(String(expense.amount ?? ""));
    setExpenseComments(expense.comments || "");
    setExpenseMessage("");
  }

  function handleCancelExpenseEdit() {
    resetExpenseEditor();
    setExpenseMessage("");
  }

  async function handleFinalizeExpenses() {
    if (activeInstallationRecord?.workflowReferenceOnly) {
      setExpenseMessage(
        "Earlier installed/downstream site expenses are read-only here.",
      );
      return;
    }

    if (!activeInstallationRecord || !expenseLedgerKey || !canEditExpenses) {
      setExpenseMessage(
        "Save this screen's Installation Details before Final Save.",
      );
      return;
    }

    if (!activeSiteExpenseApplicable) {
      setExpenseMessage(
        "Please select Yes or No for Other Installation Expenses.",
      );
      return;
    }

    if (editingExpenseId) {
      setExpenseMessage(
        "Please save the expense being edited before Final Save.",
      );
      return;
    }

    if (expenseType || expenseDate || expenseAmount || expenseComments) {
      setExpenseMessage(
        "An expense entry is still unsaved. Use Add Expense before Final Save.",
      );
      return;
    }

    if (activeSiteExpenseApplicable === "Yes" && activeSiteExpenses.length === 0) {
      setExpenseMessage(
        "Add at least one expense entry before Final Save, or select No if there are no expenses.",
      );
      return;
    }

    const now = new Date().toISOString();
    const activeRecordId = activeInstallationRecord.recordId;
    const savedCurrentCycleExpenses = getOwnedExpenses(
      activeSiteExpenses,
      activeInstallationRecord?.screenCode,
      activeInstallationRecord?.screenName,
    ).map((expense) => ({
      ...expense,
      screenCode:
        normalizeValue(expense?.screenCode) ||
        normalizeValue(activeInstallationRecord?.screenCode),
      screenName:
        normalizeValue(expense?.screenName) ||
        normalizeValue(activeInstallationRecord?.screenName),
    }));
    const historicalOwnExpenses = getHistoricalCycleExpenses(activeInstallationRecord);
    const savedExpenses = getOwnedExpenses(
      [...historicalOwnExpenses, ...savedCurrentCycleExpenses],
      activeInstallationRecord?.screenCode,
      activeInstallationRecord?.screenName,
    );
    const currentSiteTotal = getExpenseTotal(savedCurrentCycleExpenses);
    const allSiteExpensesTotal = getExpenseTotal(savedExpenses);

    const expenseSavedRecord = {
      ...activeInstallationRecord,
      otherInstallationExpensesApplicable: activeSiteExpenseApplicable,
      installationExpenses: savedExpenses,
      installationExpensesTotal: allSiteExpensesTotal,
      complexInstallationExpensesTotal:
        previousExpensesTotal + currentSiteTotal,
      installationExpensesSavedAt: now,
    };

    let persistedExpenseRecord;
    try {
      persistedExpenseRecord = await persistStage2Record(expenseSavedRecord);
    } catch (error) {
      setExpenseMessage(
        error?.message || "Other Installation Expenses could not be saved to the backend.",
      );
      return;
    }

    setStage2Records((currentRecords) =>
      currentRecords.map((record) =>
        record.recordId === activeRecordId
          ? {
              ...expenseSavedRecord,
              backendSiteId: persistedExpenseRecord.backendSiteId,
            }
          : record,
      ),
    );

    resetExpenseEditor();
    setExpenseMessage("");
    setActiveRecordId("");
    setActiveInstallationRecordId("");
    setInstallationSiteSearchTerm("");

    alert(
      "Other Installation Expenses saved. Site Information & Installation is ready for the next entry.",
    );

    window.requestAnimationFrame(() => {
      guidedScrollToElement(billingReadinessQueueRef.current);
    });
  }

  function handleEditRecord(recordId) {
    const record = stage2Records.find((item) => item.recordId === recordId);

    if (!record) {
      return;
    }

    const nextLedgerKey = getExpenseLedgerKey(record);
    const persistedExpenseState = getRecordExpenseState(record);
    const persistedExpenses = persistedExpenseState.expenses;
    const persistedApplicable =
      normalizeValue(record.otherInstallationExpensesApplicable) ||
      persistedExpenseState.applicable ||
      (persistedExpenses.length > 0 ? "Yes" : "");

    if (nextLedgerKey) {
      setExpensesBySiteId((currentExpenses) => {
        if (currentExpenses[nextLedgerKey]) {
          return currentExpenses;
        }

        return {
          ...currentExpenses,
          [nextLedgerKey]: {
            applicable: persistedApplicable,
            expenses: persistedExpenses,
          },
        };
      });
    }

    setActiveRecordId(recordId);
    setActiveInstallationRecordId(recordId);
    setInstallationSiteSearchTerm("");
    setCopySourceRecordId("");
    setCopyTargetRecordId("");
    setSelectedReadyRecordIds((currentIds) =>
      currentIds.filter((id) => id !== recordId),
    );

    window.requestAnimationFrame(() => {
      guidedScrollToElement(complexSitesRef.current);
    });
  }

  function toggleReadySelection(recordId, checked) {
    const record = filteredReadinessRows.find((row) => row.recordId === recordId);
    if (!record || !record.stage3Eligible) {
      return;
    }

    setSelectedReadyRecordIds((currentIds) =>
      checked
        ? Array.from(new Set([...currentIds, recordId]))
        : currentIds.filter((id) => id !== recordId),
    );
  }

  function handleToggleAllReadySelections(checked) {
    const visibleRecordIds = eligibleReadinessRows.map(
      (record) => record.recordId,
    );

    setSelectedReadyRecordIds((currentIds) =>
      checked
        ? Array.from(new Set([...currentIds, ...visibleRecordIds]))
        : currentIds.filter((id) => !eligibleReadinessRowIds.has(id)),
    );
  }

  function handleMoveToStage3() {
    const selectedRecords = filteredReadinessRows.filter((record) =>
      selectedReadyRecordIds.includes(record.recordId),
    );

    if (selectedRecords.length === 0) {
      alert("Select at least one eligible record to move to Stage 3.");
      return;
    }

    const eligibleRecords = selectedRecords.filter((record) => record.stage3Eligible);

    if (eligibleRecords.length === 0) {
      return;
    }

    const selectedSiteIds = new Set();
    const duplicateSelected = eligibleRecords.find((record) => {
      const screenCode = normalizeIdInput(record.screenCode);
      if (selectedSiteIds.has(screenCode)) {
        return true;
      }

      selectedSiteIds.add(screenCode);
      return false;
    });

    if (duplicateSelected) {
      alert(
        `${duplicateSelected.screenCode} appears more than once in the Stage 3 selection.`,
      );
      return;
    }

    const confirmed = window.confirm(
      "Confirm selected billing-ready sites are ready to move to Stage 3.",
    );

    if (!confirmed) {
      return;
    }

    const movedSuccessfully = onMoveToStage3(eligibleRecords);

    if (!movedSuccessfully) {
      alert("Unable to move the selected records to Stage 3.");
      return;
    }

    const movedRecordIds = new Set(eligibleRecords.map((record) => record.recordId));

    setStage2Records((currentRecords) =>
      currentRecords.filter((record) => !movedRecordIds.has(record.recordId)),
    );

    setSelectedReadyRecordIds((currentIds) =>
      currentIds.filter((recordId) => !movedRecordIds.has(recordId)),
    );

    alert(
      `${eligibleRecords.length} site record${
        eligibleRecords.length === 1 ? "" : "s"
      } moved to Stage 3.`,
    );
  }

  function handleViewReadinessDetails(recordId) {
    const record = filteredReadinessRows.find((row) => row.recordId === recordId);

    if (!record) {
      return;
    }

    setReadinessDetailsRecord(record);
  }

  function handleCloseReadinessDetails() {
    setReadinessDetailsRecord(null);
  }

  function handleDownloadExcel() {
    if (!downloadHasMatches) {
      alert("No records found for selected filters");
      return;
    }

    const exportRows = downloadFilteredRecords.map((record) => {
      const row = buildExportRow(record, expensesBySiteId, downloadStage);
      return downloadColumnHeaders.reduce((accumulator, column) => {
        accumulator[column] = row[column] ?? "";
        return accumulator;
      }, {});
    });

    const worksheet = XLSX.utils.json_to_sheet(exportRows, {
      header: downloadColumnHeaders,
    });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      "Stage 2 Billing Readiness",
    );

    const fileName = `Stage2_Billing_Readiness_${formatDateValue(new Date())}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  }

  return (
    <section
      className={`installations-page ${isMinimized ? "installations-page--minimized" : ""}`}
    >
      {readinessDetailsRecord ? (
        <div className="installations-page__modal-backdrop" role="presentation">
          <div
            className="installations-page__modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="readiness-details-title"
          >
            <div className="installations-page__modal-header">
              <div>
                <p className="installations-page__modal-eyebrow">
                  Stage 2 Billing Readiness Details
                </p>
                <h3 id="readiness-details-title">Read-only details</h3>
              </div>

              <button
                type="button"
                className="installations-page__modal-close-button"
                onClick={handleCloseReadinessDetails}
                aria-label="Close details"
                title="Close"
              >
                ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â
              </button>
            </div>

            <div className="installations-page__modal-grid">
              <div className="installations-page__modal-label">Screen Code</div>
              <div className="installations-page__modal-value">
                {readinessDetailsRecord.screenCode || "-"}
              </div>

              <div className="installations-page__modal-label">Screen Name</div>
              <div className="installations-page__modal-value">
                {readinessDetailsRecord.screenName || "-"}
              </div>

              <div className="installations-page__modal-label">Location</div>
              <div className="installations-page__modal-value">
                {readinessDetailsRecord.location || "-"}
              </div>

              <div className="installations-page__modal-label">
                Current Installation Status
              </div>
              <div className="installations-page__modal-value">
                {readinessDetailsRecord.currentStageStatus || "-"}
              </div>

              <div className="installations-page__modal-label">
                Billing Readiness Status
              </div>
              <div className="installations-page__modal-value">
                {readinessDetailsRecord.billingReadinessStatus || "-"}
              </div>

              <div className="installations-page__modal-label">Live Date</div>
              <div className="installations-page__modal-value">
                {readinessDetailsRecord.liveDate || "-"}
              </div>

              <div className="installations-page__modal-label">Billing Date</div>
              <div className="installations-page__modal-value">
                {readinessDetailsRecord.billingStartDate || "-"}
              </div>

              <div className="installations-page__modal-label">Billing Type</div>
              <div className="installations-page__modal-value">
                {readinessDetailsRecord.billingTreatment || "-"}
              </div>

              <div className="installations-page__modal-label">Billing Month</div>
              <div className="installations-page__modal-value">
                {readinessDetailsRecord.billingMonthLabel || "-"}
              </div>

              <div className="installations-page__modal-label">
                Stage 3 Eligibility
              </div>
              <div className="installations-page__modal-value">
                {formatEligibilityLabel(readinessDetailsRecord.stage3Eligible)}
              </div>

              <div className="installations-page__modal-label">Blocker Reason</div>
              <div className="installations-page__modal-value">
                {readinessDetailsRecord.blockerReason || "-"}
              </div>

              <div className="installations-page__modal-label">Remarks</div>
              <div className="installations-page__modal-value">
                {readinessDetailsRecord.remarks || "-"}
              </div>
            </div>

            <div className="installations-page__modal-footer">
              <button
                type="button"
                className="installations-page__primary-button"
                onClick={handleCloseReadinessDetails}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      ) : null}

      {activeInstallationRecord ? (
        <article
          className="installations-page__card installations__card"
          ref={siteSelectionRef}
          style={{ marginTop: "18px" }}
        >
          <div className="installations-page__installations-header">
            <h3 className="installations__card-title">
              Site Information &amp; Installation
            </h3>
            <button
              type="button"
              className="installations-page__primary-button"
              onClick={handleReturnToStage1}
              disabled={!activeRecord}
            >
              Return Selected Site to Stage 1
            </button>
          </div>

          <div className="installations__card-body">
            <div className="installations-page__field">
              <label htmlFor="installation-site-search">Search</label>
              <input
                id="installation-site-search"
                type="search"
                value={installationSiteSearchTerm}
                onChange={(event) =>
                  setInstallationSiteSearchTerm(event.target.value)
                }
                placeholder="Search Screen Code, Screen, Billing Name, Billing Code / Customer Code or Location"
              />
            </div>

            <div className="installations-page__field-grid">
              <div className="installations-page__field">
                <label>Billing Code / Customer Code</label>
                <input
                  type="text"
                  value={activeInstallationRecord.billingCode || ""}
                  readOnly
                />
              </div>
              <div className="installations-page__field">
                <label>Complex Code</label>
                <input
                  type="text"
                  value={getDisplayComplexCode(activeInstallationRecord.complexCode)}
                  readOnly
                />
              </div>
              <div className="installations-page__field">
                <label>Screen Code</label>
                <input
                  type="text"
                  value={activeInstallationRecord.screenCode || ""}
                  readOnly
                />
              </div>
              <div className="installations-page__field">
                <label>Billing Name</label>
                <input
                  type="text"
                  value={
                    activeInstallationRecord.billingName ||
                    activeInstallationRecord.screenName ||
                    ""
                  }
                  readOnly
                />
              </div>
              <div className="installations-page__field">
                <label>Location</label>
                <input
                  type="text"
                  value={activeInstallationRecord.location || ""}
                  readOnly
                />
              </div>
              <div className="installations-page__field">
                <label>State</label>
                <input
                  type="text"
                  value={activeInstallationRecord.state || ""}
                  readOnly
                />
              </div>
            </div>

            <div
              className={`installations-page__workspace installations-page__workspace--split`}
              style={{ marginTop: "18px" }}
            >
              <div
                className="installations-page__site-panel"
                ref={complexSitesRef}
              >
                <h3>
                  {normalizeValue(activeInstallationRecord.complexCode)
                    ? "Complex Sites"
                    : "Site"}
                </h3>

                <div className="installations-page__site-grid">
                  {installationVisibleRecords.length === 0 ? (
                    <div className="installations-page__empty-state">
                      No sites match the search.
                    </div>
                  ) : (
                    installationVisibleRecords.map((record) => {
                      const isActive =
                        record.recordId === activeInstallationRecord.recordId;
                      const isCopySource =
                        record.recordId === copySourceRecordId;
                      const isCopyTarget =
                        record.recordId === copyTargetRecordId;
                      const statusLabel = record.saved2B
                        ? getInstallationTableStatusLabel(
                            getInstallationTableStatusValue(record),
                          )
                        : "Installation Pending";
                      const statusValue = getInstallationTableStatusValue(record);
                      const canBeCopySource =
                        installationGroupRecords.length > 1 &&
                        record.saved2B &&
                        statusValue === "installed-completed";
                      const cardStatusLabel =
                        statusValue === "installed-completed"
                          ? "Installation Completed"
                          : statusLabel;

                      const background = isCopySource
                        ? "rgba(59, 130, 246, 0.20)"
                        : isCopyTarget
                          ? "rgba(168, 85, 247, 0.18)"
                          : !record.saved2B
                            ? "rgba(148, 163, 184, 0.08)"
                            : statusValue === "installed-completed"
                              ? "rgba(34, 197, 94, 0.12)"
                              : statusValue === "partial-installation"
                                ? "rgba(245, 158, 11, 0.12)"
                                : "rgba(239, 68, 68, 0.10)";

                      const borderColor = isCopySource
                        ? "rgba(37, 99, 235, 0.75)"
                        : isCopyTarget
                          ? "rgba(147, 51, 234, 0.75)"
                          : undefined;

                      return (
                        <button
                          key={record.recordId}
                          type="button"
                          className={`installations-page__site-btn${
                            isActive ? " installations-page__site-btn--active" : ""
                          }`}
                          style={{
                            background,
                            ...(borderColor
                              ? {
                                  borderColor,
                                  boxShadow: `0 0 0 2px ${borderColor}`,
                                }
                              : {}),
                          }}
                          onClick={() => handleInstallationSiteCardClick(record)}
                          title={
                            copySourceRecordId && !isCopySource
                              ? `Copy To ${record.screenCode || "this site"}`
                              : canBeCopySource
                                ? `Use ${record.screenCode || "this site"} as Copy From`
                                : undefined
                          }
                        >
                          <div>
                            {record.screenCode || "-"} | {record.screenName || "-"}
                          </div>

                          {isCopyTarget ? (
                            <span
                              className="installations-page__readiness-status-badge"
                              style={{
                                marginTop: "6px",
                                background: "rgba(147, 51, 234, 0.14)",
                                borderColor: "rgba(147, 51, 234, 0.45)",
                                color: "inherit",
                              }}
                            >
                              Copy To
                            </span>
                          ) : (
                            <span
                              className={`installations-page__readiness-status-badge ${getReadinessStatusTone(
                                statusLabel,
                              )}`}
                              title={
                                canBeCopySource
                                  ? isCopySource
                                    ? "Copy From selected. Click this card again to cancel."
                                    : "Click this completed card to use it as Copy From"
                                  : cardStatusLabel
                              }
                              style={{
                                marginTop: "6px",
                                ...(canBeCopySource ? { cursor: "pointer" } : {}),
                                ...(isCopySource
                                  ? {
                                      background: "rgba(37, 99, 235, 0.14)",
                                      borderColor: "rgba(37, 99, 235, 0.45)",
                                    }
                                  : {}),
                              }}
                            >
                              {isCopySource ? "Copy From" : cardStatusLabel}
                            </span>
                          )}
                        </button>
                      );
                    })
                  )}
                </div>

                {installationGroupRecords.length > 1 ? (
                  <>
                    <div className="installations-page__installation-note">
                      {completedInstallationCount} of {installationGroupRecords.length}{" "}
                      screens completed.
                    </div>
                    <div className="installations-page__installation-note">
                      {copySourceRecordId
                        ? "Copy From selected. Click another site in this Complex to copy the five installation fields."
                        : "To copy installation details, click an Installation Completed site card, then click the destination site card."}
                    </div>
                  </>
                ) : null}
              </div>

              <div className="installations-page__detail-panel">
                <h3 id="detailTitle">
                  {activeInstallationRecord.screenCode || "-"} |{" "}
                  {activeInstallationRecord.screenName || "-"}
                </h3>

                {isActiveInstallationReadOnly ? (
                  <div className="installations-page__installation-note">
                    Earlier installed/downstream site. Information is shown for
                    reference and Copy From only; this site is read-only here.
                  </div>
                ) : isBillingCommercialLocked(activeInstallationRecord) ? (
                  <div className="installations-page__installation-note">
                    Billing cycle locked after Billing Team submission. Only Dispatch Date
                    can be updated. All other installation details are view-only.
                  </div>
                ) : null}

                <div className="installations-page__meta-grid">
                  <div className="installations-page__field">
                    <label>Screen Code</label>
                    <input
                      type="text"
                      value={activeInstallationRecord.screenCode || ""}
                      readOnly
                    />
                  </div>
                  <div className="installations-page__field">
                    <label>Screen</label>
                    <input
                      type="text"
                      value={activeInstallationRecord.screenName || ""}
                      readOnly
                    />
                  </div>
                  <div className="installations-page__field">
                    <label>Location</label>
                    <input
                      type="text"
                      value={activeInstallationRecord.location || ""}
                      readOnly
                    />
                  </div>
                </div>

                <div className="installations-page__detail-grid">
                  <div className="installations-page__field">
                    <label>Dispatch Date</label>
                    <input
                      type="date"
                      max={formatDateValue(new Date())}
                      value={activeInstallationRecord.dateOfDispatch || ""}
                      disabled={isActiveInstallationReadOnly}
                      onChange={(event) =>
                        updateInstallationTableRecord(
                          activeInstallationRecord.recordId,
                          "dateOfDispatch",
                          event.target.value,
                        )
                      }
                    />
                  </div>

                  <div className="installations-page__field">
                    <label>Installation Date</label>
                    <input
                      type="date"
                      max={formatDateValue(new Date())}
                      value={activeInstallationRecord.installationDate || ""}
                      disabled={
                        isActiveInstallationReadOnly ||
                        isBillingCommercialLocked(activeInstallationRecord)
                      }
                      onChange={(event) =>
                        updateInstallationTableRecord(
                          activeInstallationRecord.recordId,
                          "installationDate",
                          event.target.value,
                        )
                      }
                    />
                  </div>

                  <div className="installations-page__field">
                    <label>Live Date</label>
                    <input
                      type="date"
                      max={formatDateValue(new Date())}
                      value={activeInstallationRecord.liveDate || ""}
                      disabled={
                        isActiveInstallationReadOnly ||
                        isBillingCommercialLocked(activeInstallationRecord)
                      }
                      onChange={(event) =>
                        updateInstallationTableRecord(
                          activeInstallationRecord.recordId,
                          "liveDate",
                          event.target.value,
                        )
                      }
                    />
                  </div>

                  <div className="installations-page__field">
                    <label>Trial Period</label>
                    <input
                      type="number"
                      min={0}
                      placeholder="Days"
                      value={activeInstallationRecord.trialPeriod || ""}
                      disabled={
                        isActiveInstallationReadOnly ||
                        isBillingCommercialLocked(activeInstallationRecord)
                      }
                      onChange={(event) =>
                        updateInstallationTableRecord(
                          activeInstallationRecord.recordId,
                          "trialPeriod",
                          event.target.value,
                        )
                      }
                    />
                  </div>

                  <div className="installations-page__field">
                    <label>Additional TP</label>
                    <input
                      type="number"
                      min={0}
                      placeholder="Days"
                      value={getNormalizedTrialPeriodExtension(
                        activeInstallationRecord,
                      )}
                      disabled={
                        isActiveInstallationReadOnly ||
                        isBillingCommercialLocked(activeInstallationRecord)
                      }
                      onChange={(event) =>
                        updateInstallationTableRecord(
                          activeInstallationRecord.recordId,
                          "trialPeriodExtension",
                          event.target.value,
                        )
                      }
                    />
                  </div>

                  <div className="installations-page__field">
                    <label>Billing Date</label>
                    <input
                      type="text"
                      value={activeInstallationRecord.billingStartDate || ""}
                      readOnly
                      placeholder="Auto"
                    />
                  </div>

                  <div className="installations-page__field">
                    <label>Current Stage Status</label>
                    {activeInstallationStatusCompleted ? (
                      <div
                        id="autoStatus"
                        className="installations-page__status-display"
                      >
                        {getInstallationTableStatusLabel("installed-completed")}
                      </div>
                    ) : (
                      <select
                        id="stageStatus"
                        value={activeInstallationStatusValue}
                        disabled={
                        isActiveInstallationReadOnly ||
                        isBillingCommercialLocked(activeInstallationRecord)
                      }
                        onChange={(event) =>
                          handleInstallationTableStatusChange(
                            activeInstallationRecord.recordId,
                            event.target.value,
                          )
                        }
                      >
                        <option value="installation-blocker">
                          Installation Blocker
                        </option>
                        <option value="partial-installation">
                          Partial Installation
                        </option>
                      </select>
                    )}
                  </div>

                  <div className="installations-page__field">
                    <label>Installation Blocker</label>
                    <input
                      type="text"
                      value={
                        activeInstallationStatusValue === "installation-blocker"
                          ? "Yes"
                          : "No"
                      }
                      readOnly
                    />
                  </div>

                  <div className="installations-page__field installations-page__field--span-2">
                    <label>Blocker Reason</label>
                    <input
                      id="blockerReason"
                      type="text"
                      value={activeInstallationRecord.blockerReason || ""}
                      placeholder={
                        activeInstallationStatusValue === "installation-blocker"
                          ? "Enter blocker reason"
                          : "Not applicable"
                      }
                      disabled={
                        isActiveInstallationReadOnly ||
                        activeInstallationStatusValue !== "installation-blocker" ||
                        isBillingCommercialLocked(activeInstallationRecord)
                      }
                      onChange={(event) =>
                        updateInstallationTableRecord(
                          activeInstallationRecord.recordId,
                          "blockerReason",
                          event.target.value,
                        )
                      }
                    />
                  </div>
                </div>

                <div className="installations-page__detail-actions">
                  <button
                    id="saveBtn"
                    type="button"
                    className="installations-page__primary-button"
                    onClick={handleSaveInstallationTableChanges}
                    disabled={isActiveInstallationReadOnly}
                    title={
                      isActiveInstallationReadOnly
                        ? "Earlier installed/downstream site is read-only"
                        : isBillingCommercialLocked(activeInstallationRecord)
                          ? "Only Dispatch Date can be updated after Billing Team submission"
                          : undefined
                    }
                  >
                    {isActiveInstallationReadOnly
                      ? "Read-only Installed Site"
                      : isBillingCommercialLocked(activeInstallationRecord)
                        ? "Save Dispatch Date"
                        : "Save Installation Details"}
                  </button>
                </div>
              </div>
            </div>


            <section
              className="installations-page__card installations-page__expense-card"
              style={{ marginTop: "18px" }}
            >
              <h3 className="installations__card-title">
                Other Installation Expenses
              </h3>

              <div className="installations__card-body">
                <div className="installations-page__installation-note">
                  {installationGroupRecords.length > 1
                    ? `${completedInstallationCount} of ${installationGroupRecords.length} screens completed. Previous sibling-site expenses are shown read-only; new expenses are saved to the current Screen Code.`
                    : "Expenses are saved to this Screen Code."}
                </div>

                <div className="installations-page__expense-grid">
                  <div className="installations-page__field">
                    <label htmlFor="expense-applicable">
                      Other Installation Expenses Applicable?
                    </label>
                    <select
                      id="expense-applicable"
                      value={activeSiteExpenseApplicable}
                      onChange={handleExpenseApplicabilityChange}
                      disabled={!canEditExpenses}
                    >
                      <option value="">Select</option>
                      <option value="Yes">Yes</option>
                      <option value="No">No</option>
                    </select>
                  </div>
                </div>

                {activeSiteExpenseApplicable === "Yes" ? (
                  <>
                    <div className="installations-page__field-grid">
                      <div className="installations-page__field">
                        <label htmlFor="expense-type">Expense Type</label>
                        <input
                          id="expense-type"
                          type="text"
                          value={expenseType}
                          onChange={(event) => setExpenseType(event.target.value)}
                          placeholder="Enter expense type"
                          disabled={!canEditExpenses}
                        />
                      </div>

                      <div className="installations-page__field">
                        <label htmlFor="expense-amount">Amount Spent (₹)</label>
                        <input
                          id="expense-amount"
                          type="number"
                          min={0}
                          value={expenseAmount}
                          onChange={(event) =>
                            setExpenseAmount(sanitizeAmount(event.target.value))
                          }
                          placeholder="Enter amount"
                          disabled={!canEditExpenses}
                        />
                      </div>

                      <div className="installations-page__field">
                        <label htmlFor="expense-date">Date of Capture</label>
                        <input
                          id="expense-date"
                          type="date"
                          max={formatDateValue(new Date())}
                          value={expenseDate}
                          onChange={(event) => setExpenseDate(event.target.value)}
                          disabled={!canEditExpenses}
                        />
                      </div>

                      <div className="installations-page__field">
                        <label htmlFor="expense-comments">Remarks</label>
                        <textarea
                          id="expense-comments"
                          value={expenseComments}
                          onChange={(event) => setExpenseComments(event.target.value)}
                          placeholder="Enter remarks"
                          disabled={!canEditExpenses}
                        />
                      </div>
                    </div>

                    <div className="installations-page__detail-actions">
                      <button
                        type="button"
                        className="installations-page__primary-button"
                        onClick={handleSaveExpense}
                        disabled={!canEditExpenses}
                      >
                        {editingExpenseId ? "Save Expense Changes" : "Add Expense"}
                      </button>
                      {editingExpenseId ? (
                        <button
                          type="button"
                          className="installations-page__edit-button"
                          onClick={handleCancelExpenseEdit}
                        >
                          Cancel Edit
                        </button>
                      ) : null}
                    </div>
                  </>
                ) : null}

                {previousComplexExpenses.length > 0 ||
                activeSiteExpenses.length > 0 ? (
                  <div className="installations-page__expense-history">
                    <h4>Expense Ledger</h4>
                    <div className="installations-page__search-table-wrapper">
                      <table className="installations-page__search-table">
                        <thead>
                          <tr>
                            <th>Screen</th>
                            <th>Expense Type</th>
                            <th>Amount Spent</th>
                            <th>Date of Capture</th>
                            <th>Remarks</th>
                            <th>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {previousComplexExpenses.map((expense) => (
                            <tr key={`previous-${expense.expenseId}`}>
                              <td>
                                {expense.screenCode || "-"}
                                {expense.screenName
                                  ? ` / ${expense.screenName}`
                                  : ""}
                              </td>
                              <td>{expense.expenseType || "-"}</td>
                              <td>₹{formatExpenseAmount(expense.amount)}</td>
                              <td>{formatExpenseDateDisplay(expense.expenseDate)}</td>
                              <td>{expense.comments || "-"}</td>
                              <td>Read-only</td>
                            </tr>
                          ))}

                          {activeSiteExpenses.map((expense) => (
                            <tr key={expense.expenseId}>
                              <td>
                                {expense.screenCode ||
                                  activeInstallationRecord?.screenCode ||
                                  "-"}
                                {(expense.screenName ||
                                  activeInstallationRecord?.screenName)
                                  ? ` / ${
                                      expense.screenName ||
                                      activeInstallationRecord?.screenName
                                    }`
                                  : ""}
                              </td>
                              <td>{expense.expenseType || "-"}</td>
                              <td>₹{formatExpenseAmount(expense.amount)}</td>
                              <td>{formatExpenseDateDisplay(expense.expenseDate)}</td>
                              <td>{expense.comments || "-"}</td>
                              <td>
                                {canEditExpenses ? (
                                  <button
                                    type="button"
                                    className="installations-page__edit-button"
                                    onClick={() =>
                                      handleEditExpense(expense.expenseId)
                                    }
                                  >
                                    Edit
                                  </button>
                                ) : (
                                  "Read-only"
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    <div
                      className="installations-page__expense-total"
                      style={{
                        display: "grid",
                        gridTemplateColumns: "minmax(0, 1fr) auto",
                        gap: "8px 18px",
                        alignItems: "center",
                        marginTop: "10px",
                        padding: "10px 12px",
                      }}
                    >
                      <span>Previous Expenses</span>
                      <strong>₹{formatExpenseAmount(previousExpensesTotal)}</strong>

                      <span>Current Site Expenses</span>
                      <strong>₹{formatExpenseAmount(currentSiteExpensesTotal)}</strong>

                      <span style={{ fontWeight: 800 }}>Total</span>
                      <strong style={{ fontSize: "1.05rem" }}>
                        ₹{formatExpenseAmount(combinedExpensesTotal)}
                      </strong>
                    </div>
                  </div>
                ) : null}

                {!hasAnySavedInstallationDetails ? (
                  <div className="installations-page__search-empty">
                    Save at least one screen's Installation Details to enable
                    Other Installation Expenses.
                  </div>
                ) : null}

                <div id="stageMessage" className="installations-page__message">
                  {expenseMessage}
                </div>

                <div
                  className="installations-page__detail-actions"
                  style={{ justifyContent: "flex-end", marginTop: "14px" }}
                >
                  <button
                    type="button"
                    className="installations-page__primary-button"
                    onClick={handleFinalizeExpenses}
                    disabled={!canEditExpenses}
                  >
                    Final Save
                  </button>
                </div>
              </div>
            </section>
          </div>
        </article>
      ) : (
        <article
          className="installations-page__card installations__card"
          ref={siteSelectionRef}
          style={{ marginTop: "18px" }}
        >
          <div className="installations-page__installations-header">
            <h3 className="installations__card-title">Site Information &amp; Installation</h3>
            <button type="button" className="installations-page__primary-button" disabled>
              Return Selected Site to Stage 1
            </button>
          </div>

          <div className="installations__card-body">
            <div className="installations-page__field">
              <label>Search</label>
              <input
                type="search"
                value=""
                placeholder="Search Screen Code, Screen, Billing Name, Billing Code / Customer Code or Location"
                disabled
                readOnly
              />
            </div>

            <div className="installations-page__field-grid">
              {["Billing Code / Customer Code", "Complex Code", "Screen Code", "Billing Name", "Location", "State"].map((label) => (
                <div className="installations-page__field" key={label}>
                  <label>{label}</label>
                  <input type="text" value="" readOnly disabled />
                </div>
              ))}
            </div>

            <div className="installations-page__workspace installations-page__workspace--split" style={{ marginTop: "18px" }}>
              <div className="installations-page__site-panel">
                <h3>Sites</h3>
                <div className="installations-page__empty-state">No site selected.</div>
              </div>

              <div className="installations-page__detail-panel">
                <h3>Selected Site</h3>
                <div className="installations-page__detail-grid">
                  {["Screen Code", "Screen", "Location", "Dispatch Date", "Installation Date", "Live Date", "Trial Period", "Additional TP", "Billing Date", "Current Stage Status", "Installation Blocker", "Blocker Reason"].map((label) => (
                    <div className="installations-page__field" key={label}>
                      <label>{label}</label>
                      <input type="text" value="" readOnly disabled />
                    </div>
                  ))}
                </div>
                <div className="installations-page__detail-actions">
                  <button type="button" className="installations-page__primary-button" disabled>
                    Save Installation Details
                  </button>
                </div>
              </div>
            </div>

            <section className="installations-page__card installations-page__expense-card" style={{ marginTop: "18px" }}>
              <h3 className="installations__card-title">Other Installation Expenses</h3>
              <div className="installations__card-body">
                <div className="installations-page__field">
                  <label>Other Installation Expenses Applicable?</label>
                  <select value="" disabled>
                    <option value="">Select</option>
                  </select>
                </div>
                <div className="installations-page__detail-actions" style={{ justifyContent: "flex-end", marginTop: "14px" }}>
                  <button type="button" className="installations-page__primary-button" disabled>
                    Final Save
                  </button>
                </div>
              </div>
            </section>
          </div>
        </article>
      )}

      <section
        className="installations-page__billing-readiness"
        ref={billingReadinessQueueRef}
      >
        <div className="installations-page__billing-readiness-card">
          <div className="installations-page__billing-readiness-header">
            <div>
              <h2>Billing Readiness Queue</h2>
            </div>
            <div style={{ display: "flex", alignItems: "end", gap: 10, marginLeft: "auto" }}>
              <button
                type="button"
                className="installations-page__primary-button"
                onClick={handleMoveToStage3}
                disabled={!hasSelectedReadyRecords}
              >
                Move Selected to Stage 3
              </button>
            </div>
          </div>

          <div className="installations-page__stage-filter">
            <div className="installations-page__field installations-page__stage-filter-field">
              <label htmlFor="readiness-search-input">Search</label>
              <input
                id="readiness-search-input"
                type="search"
                value={readinessSearchTerm}
                onChange={(event) => setReadinessSearchTerm(event.target.value)}
                placeholder="Search Screen Code / Screen / Location"
              />
            </div>

            <div className="installations-page__field installations-page__stage-filter-field">
              <label htmlFor="readiness-stage-filter">Stage</label>
              <select
                id="readiness-stage-filter"
                value={readinessStageFilter}
                onChange={(event) => setReadinessStageFilter(event.target.value)}
              >
                {readinessStageFilterOptions.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>
            </div>

            <div className="installations-page__field">
              <label htmlFor="readiness-from-date">From Date</label>
              <input
                id="readiness-from-date"
                type="date"
                value={readinessFromDate}
                onChange={(event) => setReadinessFromDate(event.target.value)}
              />
            </div>

            <div className="installations-page__field">
              <label htmlFor="readiness-to-date">To Date</label>
              <input
                id="readiness-to-date"
                type="date"
                value={readinessToDate}
                onChange={(event) => setReadinessToDate(event.target.value)}
              />
            </div>

            <div className="installations-page__download-wrap">
              <button
                type="button"
                className="installations-page__primary-button"
                onClick={handleDownloadExcel}
                disabled={!downloadHasMatches}
              >
                Download Excel
              </button>
            </div>
          </div>

          <div className="installations-page__readiness-table-wrapper">
            <table className="installations-page__readiness-table" style={{ tableLayout: "fixed", width: "100%" }}>
              <colgroup>
                <col className="installations-page__readiness-col-select" style={{ width: "52px" }} />
                <col className="installations-page__readiness-col-site" style={{ width: "112px" }} />
                <col className="installations-page__readiness-col-screen" />
                <col className="installations-page__readiness-col-location" />
                <col className="installations-page__readiness-col-live-date" style={{ width: "108px" }} />
                <col className="installations-page__readiness-col-billing-date" style={{ width: "116px" }} />
                <col className="installations-page__readiness-col-status" style={{ width: "168px" }} />
                <col className="installations-page__readiness-col-action" style={{ width: "84px" }} />
              </colgroup>
              <thead>
                <tr>
                  <th className="installations-page__readiness-table-select-header" style={{ textAlign: "center" }}>
                    <input
                      ref={readinessSelectAllRef}
                      type="checkbox"
                      checked={areAllVisibleReadyRecordsSelected}
                      onChange={(event) =>
                        handleToggleAllReadySelections(event.target.checked)
                      }
                      aria-label="Select all Stage 2 billing readiness records"
                      title="Select all"
                      disabled={!hasVisibleReadyRecords}
                    />
                  </th>
                  <SortableHeader label="Screen Code" sortKey="screenCode" activeKey={readinessSortKey} direction={readinessSortOrder} onSort={handleReadinessColumnSort} style={{ textAlign: "left" }} />
                  <SortableHeader label="Screen" sortKey="screenName" activeKey={readinessSortKey} direction={readinessSortOrder} onSort={handleReadinessColumnSort} style={{ textAlign: "left" }} />
                  <SortableHeader label="Location" sortKey="location" activeKey={readinessSortKey} direction={readinessSortOrder} onSort={handleReadinessColumnSort} style={{ textAlign: "left" }} />
                  <SortableHeader label="Live Date" sortKey="liveDate" activeKey={readinessSortKey} direction={readinessSortOrder} onSort={handleReadinessColumnSort} className="installations-page__readiness-table-live-date-header" style={{ textAlign: "left" }} />
                  <SortableHeader label="Billing Date" sortKey="billingStartDate" activeKey={readinessSortKey} direction={readinessSortOrder} onSort={handleReadinessColumnSort} className="installations-page__readiness-table-billing-date-header" style={{ textAlign: "left" }} />
                  <SortableHeader label="Billing Readiness" sortKey="billingReadiness" activeKey={readinessSortKey} direction={readinessSortOrder} onSort={handleReadinessColumnSort} style={{ textAlign: "center" }} />
                  <th className="installations-page__readiness-table-action-header" style={{ textAlign: "center" }}>
                    Action
                  </th>
                </tr>
              </thead>
              <tbody>
                {filteredReadinessRows.length === 0 ? (
                  <tr>
                    <td
                      className="installations-page__billing-readiness-empty-row"
                      colSpan={8}
                    >
                      No billing readiness records match the selected stage.
                    </td>
                  </tr>
                ) : (
                  filteredReadinessRows.map((record) => {
                    const isSelected = selectedReadyRecordIds.includes(record.recordId);
                    const screenValue = getDisplayValue(record.screenName);
                    const locationValue = getDisplayValue(record.location);
                    const readinessValue = formatReadinessBadgeLabel(
                      getDisplayValue(
                        record.billingReadinessStatus || record.currentStageStatus,
                      ),
                    );

                    return (
                      <tr
                        key={record.recordId}
                        className={
                          record.billingReadinessStatus === "Future Billing"
                            ? "installations-page__readiness-row--future"
                            : ""
                        }
                      >
                        <td className="installations-page__readiness-table-select-cell" style={{ textAlign: "center", verticalAlign: "middle" }}>
                          <input
                            type="checkbox"
                            checked={isSelected}
                            onChange={(event) =>
                              toggleReadySelection(record.recordId, event.target.checked)
                            }
                            aria-label={`Select ${record.screenCode}`}
                            disabled={!record.stage3Eligible}
                          />
                        </td>
                        <td title={record.screenCode} style={{ textAlign: "left", verticalAlign: "middle" }}>{record.screenCode}</td>
                        <td
                          className="installations-page__readiness-screen-cell"
                          title={screenValue}
                          style={{ textAlign: "left", verticalAlign: "middle" }}
                        >
                          {screenValue}
                        </td>
                        <td
                          className="installations-page__readiness-location-cell"
                          title={locationValue}
                          style={{ textAlign: "left", verticalAlign: "middle" }}
                        >
                          {locationValue}
                        </td>
                        <td
                          className="installations-page__readiness-table-live-date-cell"
                          title={record.liveDate || "-"}
                          style={{ textAlign: "left", verticalAlign: "middle" }}
                        >
                          {record.liveDate || "-"}
                        </td>
                        <td
                          className="installations-page__readiness-table-billing-date-cell"
                          title={record.billingStartDate || "-"}
                          style={{ textAlign: "left", verticalAlign: "middle" }}
                        >
                          {record.billingStartDate || "-"}
                        </td>
                        <td
                          className="installations-page__readiness-status-cell"
                          title={readinessValue}
                          style={{ textAlign: "center", verticalAlign: "middle" }}
                        >
                          <div className="installations-page__readiness-status-stack" style={{ display: "flex", justifyContent: "center", alignItems: "center" }}>
                            <span
                              className={`installations-page__readiness-status-badge ${getReadinessStatusTone(
                                record.billingReadinessStatus || record.currentStageStatus,
                              )}`}
                              style={{
                                maxWidth: "100%",
                                minWidth: 0,
                                justifyContent: "center",
                                textAlign: "center",
                                boxSizing: "border-box",
                              }}
                            >
                              {readinessValue}
                            </span>
                          </div>
                        </td>
                        <td className="installations-page__readiness-table-action-cell" style={{ textAlign: "center", verticalAlign: "middle" }}>
                          <button
                            type="button"
                            className="installations-page__edit-button"
                            onClick={() => handleEditRecord(record.recordId)}
                            aria-label={`Edit ${record.screenCode}`}
                          >
                            Edit
                          </button>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      </section>

    </section>
  );
}

export default Installations;
