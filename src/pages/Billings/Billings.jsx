import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import "./Billings.css";
import { isValidDateValue, isValidFourDigitYear } from "../../utils/dateValidation";
import { getDisplayComplexCode } from "../../utils/idValidation";
import { sanitizeAmount } from "../../utils/numberInput";
import {
  billingApiRequest,
  ensureAllocationBackendMigration,
  getStage3BillingSummary,
} from "../../utils/billingApi";
import {
  getBillingAsOfDate,
  getLatestNormalBillingMonthStart,
  isNormalBillingMonthAvailable,
} from "../../utils/billingPeriodAvailability";

const GUIDED_SCROLL_DURATION_MS = 800;
const GUIDED_SCROLL_TOP_OFFSET_PX = 24;

function formatPriceChangeDateDisplay(value) {
  const match = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? `${match[3]}-${match[2]}-${match[1]}` : "";
}

function maskPriceChangeDateInput(value) {
  const digits = String(value || "").replace(/\D/g, "").slice(0, 8);
  if (digits.length <= 2) return digits;
  if (digits.length <= 4) return `${digits.slice(0, 2)}-${digits.slice(2)}`;
  return `${digits.slice(0, 2)}-${digits.slice(2, 4)}-${digits.slice(4)}`;
}

function getCanonicalPriceChangeDate(displayValue) {
  const match = String(displayValue || "").match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!match) return "";

  const canonical = `${match[3]}-${match[2]}-${match[1]}`;
  return isValidDateValue(canonical) ? canonical : "";
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

const workspaceModes = {
  EXPENSES: "Other Expenses",
  OTF: "OTF",
  BILLING: "Billing",
  FOC: "FoC",
};

const billingSections = {
  FIRST_TIME: "first-time",
  REALLOCATION: "reallocation",
  RECURRING: "recurring",
  RECORDS: "records",
  CLOSURE: "closure",
};

const humanErrorCorrectionFields = [
  { key: "subscriptionFee", label: "Subscription Fee" },
  { key: "subscriptionMode", label: "Subscription Mode" },
  { key: "incentiveBeneficiary", label: "Incentive Beneficiary" },
  { key: "otfAmount", label: "OTF Amount" },
];

const billingAuditFieldLabels = {
  invoiceNumber: "Invoice Number",
  invoiceDate: "Invoice Date",
  invoiceAmountBeforeGST: "Invoice Amount Before GST",
  invoiceStatus: "Invoice Status",
  paymentStatus: "Payment Status",
  paymentReceivedDate: "Payment Received Date",
  paymentRemarks: "Payment Remarks",
  billingVerificationStatus: "Billing Verification Status",
  stage: "Stage",
  correctionFields: "Correction Fields",
};

const rowsPerPageOptions = [10, 25, 50, 100];

function normalizeStage3BillingSummary(summary) {
  if (!summary || typeof summary !== "object") {
    return null;
  }

  if (
    summary.data &&
    typeof summary.data === "object" &&
    !Array.isArray(summary.data) &&
    !("billingRecords" in summary)
  ) {
    return summary.data;
  }

  return summary;
}

function formatSummaryCount(summarySection) {
  const count = Number(summarySection?.count);
  return Number.isFinite(count) ? count : "Unavailable";
}

function normalizeValue(value) {
  return String(value || "").trim();
}

function resolveBillingRecordSite(billingRecord, candidates = []) {
  if (!billingRecord) return null;

  const backendSiteId = normalizeValue(
    billingRecord.backendSiteId || billingRecord.siteBackendId,
  );
  const screenCode = normalizeValue(
    billingRecord.screenCode || billingRecord.siteId || billingRecord.siteScope,
  ).toUpperCase();

  return (
    candidates.find(
      (candidate) => {
        const candidateBackendId = normalizeValue(
          candidate?.backendSiteId ||
            candidate?.id ||
            candidate?.stage3Data?.backendSiteId ||
            candidate?.stage2Data?.backendSiteId ||
            candidate?.stage1Data?.backendSiteId,
        );
        return backendSiteId && candidateBackendId === backendSiteId;
      },
    ) ||
    candidates.find(
      (candidate) => {
        const candidateScreenCode = normalizeValue(
          candidate?.screenCode ||
            candidate?.siteId ||
            candidate?.stage3Data?.screenCode ||
            candidate?.stage2Data?.screenCode ||
            candidate?.stage1Data?.screenCode,
        ).toUpperCase();
        return screenCode && candidateScreenCode === screenCode;
      },
    ) ||
    null
  );
}

function getOwnedInstallationExpenses(record) {
  const screenCode = normalizeValue(record?.screenCode).toUpperCase();

  return (Array.isArray(record?.installationExpenses)
    ? record.installationExpenses
    : [])
    .filter((expense) => {
      const expenseScreenCode = normalizeValue(expense?.screenCode).toUpperCase();
      return !expenseScreenCode || expenseScreenCode === screenCode;
    })
    .map((expense) => ({
      ...expense,
      screenCode: screenCode || normalizeValue(expense?.screenCode),
      screenName: normalizeValue(expense?.screenName) || normalizeValue(record?.screenName),
    }));
}


function normalizeSubscriptionModeLabel(value) {
  const normalized = normalizeValue(value)
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (["annual", "yearly", "one year"].includes(normalized)) return "Annual";
  if (["half year", "half yearly", "semi annual"].includes(normalized)) return "Half-Yearly";
  return "Monthly";
}

function getDefaultAlternativeSubscriptionMode(currentMode) {
  const current = normalizeSubscriptionModeLabel(currentMode);
  if (current === "Annual") return "Monthly";
  if (current === "Half-Yearly") return "Monthly";
  return "Half-Yearly";
}

function isAllocationModelBillingRecord(record) {
  return normalizeValue(record?.pricingMethod).toLowerCase() === "allocation model";
}

function getAllocationIdentity(value, fallback = "") {
  const explicitRecordChain =
    normalizeValue(value?.allocationChainNetwork) ||
    normalizeValue(value?.chainNetwork);

  const explicitPurpose =
    normalizeValue(value?.allocationPurpose) ||
    normalizeValue(value?.purpose) ||
    normalizeValue(value?.allocationSettingsSnapshot?.allocationPurpose) ||
    normalizeValue(value?.allocationSettingsSnapshot?.purpose);

  const rawIdentity =
    explicitRecordChain ||
    normalizeValue(value?.chainName) ||
    normalizeValue(value?.networkName) ||
    normalizeValue(value?.allocationSettingsSnapshot?.chainName) ||
    normalizeValue(value?.allocationSettingsSnapshot?.modelName) ||
    normalizeValue(value?.name) ||
    normalizeValue(value?.modelName) ||
    normalizeValue(fallback);

  let chainName = rawIdentity;
  let allocationPurpose = explicitPurpose;

  // Allocation display names may have been saved historically as:
  // "UYLJK Stage 3" or "UYLJK - Stage 3".
  // Strip the Stage 3 suffix from the Chain identity even when the
  // allocationPurpose field has already been saved separately.
  if (!explicitRecordChain) {
    const stage3Match = rawIdentity.match(
      /^(.*?)\s*(?:[-–—·:]\s*)?Stage\s*3$/i,
    );
    const normalizedPurpose = allocationPurpose
      .toLowerCase()
      .replace(/[-_]/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    if (
      stage3Match?.[1] &&
      (!allocationPurpose ||
        normalizedPurpose === "stage 3" ||
        normalizedPurpose === "stage3")
    ) {
      chainName = normalizeValue(stage3Match[1]);
      allocationPurpose = allocationPurpose || "Stage 3";
    }
  }

  const displayName = [chainName, allocationPurpose].filter(Boolean).join(" ");

  return {
    chainName,
    allocationPurpose,
    displayName: displayName || rawIdentity || normalizeValue(fallback),
  };
}

function getAllocationChainMatchKey(value) {
  return normalizeValue(value)
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function getAllocationChainName(value, fallback = "") {
  return getAllocationIdentity(value, fallback).chainName;
}

function getAllocationPurpose(value, fallback = "") {
  return (
    getAllocationIdentity(value, fallback).allocationPurpose ||
    normalizeValue(fallback)
  );
}

function getAllocationDisplayName(value, fallback = "") {
  return getAllocationIdentity(value, fallback).displayName;
}

function getAllocationModelBillingGroupKey(record) {
  return (
    normalizeValue(record?.stage1GroupId) ||
    normalizeValue(record?.billingCode) ||
    normalizeValue(record?.complexCode) ||
    normalizeValue(record?.recordId)
  );
}

function normalizeAllocationModel(model, fallbackName = "Allocation Model") {
  if (!model || typeof model !== "object") return null;

  const allocationIdentity = getAllocationIdentity(model, fallbackName);

  let slabs = Array.isArray(model.slabs) ? model.slabs : [];
  if (slabs.length === 0) {
    const legacy = [];
    if (normalizeValue(model.upTo3PlanFee)) {
      legacy.push({
        slabId: "up-to-3",
        label: "Up to 3",
        min: 1,
        max: 3,
        planFee: normalizeValue(model.upTo3PlanFee),
      });
    }
    if (normalizeValue(model.moreThan3PlanFee)) {
      legacy.push({
        slabId: "more-than-3",
        label: "More Than 3",
        min: 4,
        max: null,
        planFee: normalizeValue(model.moreThan3PlanFee),
      });
    }
    slabs = legacy;
  }

  return {
    modelId:
      normalizeValue(model.modelId) ||
      normalizeValue(model.id) ||
      `allocation-model-${normalizeValue(model.effectiveFrom) || "legacy"}`,
    chainName: allocationIdentity.chainName,
    allocationPurpose: allocationIdentity.allocationPurpose,
    name: allocationIdentity.displayName,
    status: normalizeValue(model.status),
    effectiveFrom: normalizeValue(model.effectiveFrom),
    effectiveTo: normalizeValue(model.effectiveTo),
    otfPricePerDevice: normalizeValue(model.otfPricePerDevice),
    slabs: slabs.map((slab, index) => ({
      slabId: normalizeValue(slab?.slabId) || `slab-${index + 1}`,
      label: normalizeValue(slab?.label) || `Plan ${index + 1}`,
      min: Number(slab?.min) || 1,
      max:
        slab?.max === null || slab?.max === undefined || slab?.max === ""
          ? null
          : Number(slab.max),
      planFee: normalizeValue(slab?.planFee),
    })),
  };
}

function getConfiguredAllocationModelsForRecord(record, configuredModels = []) {
  const recordChain = getAllocationChainMatchKey(
    getAllocationChainName(record),
  );

  return (Array.isArray(configuredModels) ? configuredModels : []).filter((model) => {
    const modelChain = getAllocationChainMatchKey(
      getAllocationChainName(model),
    );
    const status = normalizeValue(model.status).toLowerCase();
    const purpose = getAllocationPurpose(model)
      .toLowerCase()
      .replace(/[-_]/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    const sameChain = !recordChain || modelChain === recordChain;
    const activeOrLegacy = !status || status === "active";
    const stage3Purpose =
      purpose === "stage 3" ||
      purpose === "stage3";

    return sameChain && activeOrLegacy && stage3Purpose;
  });
}

function findAllocationModelSlab(model, deviceCount) {
  const count = Math.max(1, Number(deviceCount) || 1);
  const slabs = Array.isArray(model?.slabs) ? model.slabs : [];

  return (
    slabs.find((slab) => {
      const min = Number(slab?.min) || 1;
      const max =
        slab?.max === null || slab?.max === undefined || slab?.max === ""
          ? null
          : Number(slab.max);
      return count >= min && (max === null || count <= max);
    }) || null
  );
}

function formatAllocationModelSummary(model) {
  const normalized = normalizeAllocationModel(model);
  if (!normalized || normalized.slabs.length === 0) return "-";
  return normalized.slabs
    .map((slab) => `${slab.label}: ₹${normalizeValue(slab.planFee) || "0"}`)
    .join(" | ");
}

function isFoCOriginRecord(record) {
  if (!record) return false;
  if (record.focOrigin === true || record.foCOrigin === true) return true;

  const directValues = [
    record.foc,
    record.foC,
    record.focApplicable,
    record.isFoc,
    record.isFoC,
    record.freeOfCharge,
    record.focStatus,
  ]
    .map((value) => normalizeValue(value).toLowerCase())
    .filter(Boolean);

  if (directValues.some((value) => ["yes", "true", "1", "foc", "free of charge"].includes(value))) {
    return true;
  }

  const statusValues = [
    record.originalCommercialStatus,
    record.commercialStatus,
    record.billingMode,
    record.pricingStatus,
  ]
    .map((value) => normalizeValue(value).toLowerCase())
    .filter(Boolean);

  return statusValues.some((value) => value === "foc" || value === "free of charge");
}

function isSubscriptionApplicableRecord(record) {
  if (!record || isFoCOriginRecord(record)) {
    return false;
  }

  const explicitApplicability = [
    record.subscriptionApplicable,
    record.subscriptionEligible,
    record.subscriptionBillingApplicable,
  ]
    .map((value) => normalizeValue(value).toLowerCase())
    .find(Boolean);

  if (["no", "false", "0", "not applicable", "n/a"].includes(explicitApplicability)) {
    return false;
  }

  if (normalizeValue(record.commercialApplicable).toLowerCase() === "false") {
    return false;
  }

  const subscriptionType = normalizeValue(record.subscriptionType)
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const subscriptionMode = normalizeValue(record.subscriptionMode)
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  // OTF-only and otherwise non-subscription screens may share a complex with
  // subscription screens, but must not enter subscription billing.
  if (
    !subscriptionType ||
    ["no", "none", "not applicable", "n/a", "na"].includes(subscriptionType) ||
    !subscriptionMode
  ) {
    return false;
  }

  if (!getBillingPeriodFrequency(record)) {
    return false;
  }

  const configuredAmounts = [
    record.subscriptionFee,
    ...(Array.isArray(record.pricingGroups)
      ? record.pricingGroups.map((group) => group?.subscriptionFee)
      : []),
    ...(Array.isArray(record.allocationRows)
      ? record.allocationRows.map((row) => row?.planFee)
      : []),
    ...Object.values(
      record.monthlyBillingAmounts && typeof record.monthlyBillingAmounts === "object"
        ? record.monthlyBillingAmounts
        : {},
    ),
  ];

  return configuredAmounts.some((value) => {
    const amount = parseAmountValue(value);
    return amount !== null && amount > 0;
  });
}

function isFoCBillingDateEligible(record) {
  const billingDate = parseLocalDateValue(record?.billingStartDate);
  if (!billingDate) return false;
  const today = new Date();
  const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return billingDate <= todayStart;
}

function getFoCCurrentCommercialStatus(record) {
  return normalizeValue(record?.focCurrentCommercialStatus) || "FoC";
}

function formatAuditValue(value) {
  if (Array.isArray(value)) {
    return value.length ? value.join(", ") : "-";
  }

  if (value === null || value === undefined || value === "") {
    return "-";
  }

  return String(value);
}

function getBillingAuditFieldLabel(field) {
  return billingAuditFieldLabels[field] || field || "-";
}

function createBillingAuditEntry({
  event,
  field = "",
  previousValue = "",
  newValue = "",
  remarks = "",
  timestamp = new Date().toISOString(),
}) {
  return {
    auditId: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    timestamp,
    event: normalizeValue(event) || "Updated",
    field: normalizeValue(field),
    previousValue: formatAuditValue(previousValue),
    newValue: formatAuditValue(newValue),
    remarks: normalizeValue(remarks),
  };
}

function getCorrectionDisplayLabel(record) {
  if (
    normalizeValue(record?.correctionType) === "Human Error" ||
    normalizeValue(record?.workflowType) === "RECORD_CORRECTION"
  ) {
    return "Record Correction";
  }

  return "";
}

function parseDate(value) {
  if (!isValidDateValue(value)) {
    return null;
  }

  const parsed = new Date(`${value}T00:00:00`);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function parseLocalDateValue(value) {
  if (!isValidDateValue(value)) {
    return null;
  }

  const [year, month, day] = String(value).split("-").map(Number);
  const parsed = new Date(year, month - 1, day);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}



function formatBillingDateDisplay(value) {
  const normalized = normalizeValue(value);

  if (!normalized) {
    return "-";
  }

  const date = parseLocalDateValue(normalized);

  if (!date) {
    return normalized;
  }

  return date.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });
}

function isFutureLocalDateValue(value) {
  const parsed = parseLocalDateValue(value);

  if (!parsed) {
    return false;
  }

  const today = new Date();
  const todayStart = new Date(
    today.getFullYear(),
    today.getMonth(),
    today.getDate(),
  );

  return parsed > todayStart;
}

function isPaymentReceivedStatus(value) {
  return ["paid", "payment received", "received", "completed", "complete"].includes(
    normalizeValue(value).toLowerCase(),
  );
}


function formatLocalDateValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

function addCalendarMonths(date, months) {
  const targetYear = date.getFullYear();
  const targetMonth = date.getMonth() + months;
  const targetMonthLastDay = new Date(targetYear, targetMonth + 1, 0).getDate();
  const targetDay = Math.min(date.getDate(), targetMonthLastDay);

  return new Date(targetYear, targetMonth, targetDay);
}

function addCalendarYears(date, years) {
  return addCalendarMonths(date, years * 12);
}

function getBillingPeriodFrequency(record) {
  const normalizedMode = normalizeValue(record?.subscriptionMode)
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (
    normalizedMode === "yearly" ||
    normalizedMode === "annual" ||
    normalizedMode === "one year"
  ) {
    return "year";
  }

  if (
    normalizedMode === "half year" ||
    normalizedMode === "half yearly" ||
    normalizedMode === "half yearly" ||
    normalizedMode === "semi annual" ||
    normalizedMode === "semi annual"
  ) {
    return "half-year";
  }

  if (normalizedMode === "quarterly" || normalizedMode === "quarter") {
    return "quarter";
  }

  if (normalizedMode === "monthly" || normalizedMode === "month" || normalizedMode === "") {
    return "month";
  }

  return null;
}

function getCurrentMonthEndValue() {
  const today = new Date();
  const currentMonthEnd = new Date(
    today.getFullYear(),
    today.getMonth() + 1,
    0,
  );

  return formatLocalDateValue(currentMonthEnd);
}

function isPastBillingMonth(dateValue) {
  const parsed = parseLocalDateValue(dateValue);

  if (!parsed) {
    return false;
  }

  const today = new Date();
  const billingMonthStart = new Date(
    parsed.getFullYear(),
    parsed.getMonth(),
    1,
  );
  const currentMonthStart = new Date(
    today.getFullYear(),
    today.getMonth(),
    1,
  );

  return billingMonthStart < currentMonthStart;
}

function calculateBillingPeriodTo(fromValue, record) {
  const fromDate = parseLocalDateValue(fromValue);

  if (!fromDate) {
    return "";
  }

  const frequency = getBillingPeriodFrequency(record);
  if (!frequency) {
    return "";
  }

  if (frequency === "month") {
    const normalizedSubscriptionType = normalizeValue(record?.subscriptionType)
      .toLowerCase();
    const isVariableMonthly = normalizedSubscriptionType === "variable";

    if (isVariableMonthly && isPastBillingMonth(fromValue)) {
      const today = new Date();
      const previousMonthEnd = new Date(
        today.getFullYear(),
        today.getMonth(),
        0,
      );

      return formatLocalDateValue(previousMonthEnd);
    }

    if (isPastBillingMonth(fromValue)) {
      return getCurrentMonthEndValue();
    }

    return formatLocalDateValue(
      new Date(fromDate.getFullYear(), fromDate.getMonth() + 1, 0),
    );
  }

  if (frequency === "quarter") {
    const nextDate = addCalendarMonths(fromDate, 3);
    nextDate.setDate(nextDate.getDate() - 1);
    return formatLocalDateValue(nextDate);
  }

  if (frequency === "half-year") {
    const nextDate = addCalendarMonths(fromDate, 6);
    nextDate.setDate(nextDate.getDate() - 1);
    return formatLocalDateValue(nextDate);
  }

  if (frequency === "year") {
    const nextDate = addCalendarYears(fromDate, 1);
    nextDate.setDate(nextDate.getDate() - 1);
    return formatLocalDateValue(nextDate);
  }

  return "";
}

function getBillingMonthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function getBillingMonthLabel(date) {
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    year: "numeric",
  }).format(date);
}

function getFinancialYearLabelFromDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
    return "";
  }

  const startYear =
    date.getMonth() >= 3 ? date.getFullYear() : date.getFullYear() - 1;
  const endYearShort = String(startYear + 1).slice(-2);

  return `${startYear}-${endYearShort}`;
}

function getCurrentFinancialYearLabel() {
  return getFinancialYearLabelFromDate(new Date());
}

function buildMonthlyBillingRows(
  record,
  fixedMonthlyFeeOverride,
  asOfDate = new Date(),
) {
  const billingStartDate = parseLocalDateValue(record?.billingStartDate);

  if (!billingStartDate) {
    return [];
  }

  const subscriptionType = normalizeValue(record?.subscriptionType).toLowerCase();
  const subscriptionMode = normalizeValue(record?.subscriptionMode)
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!["monthly", "month", ""].includes(subscriptionMode)) {
    return [];
  }

  const isVariable = subscriptionType === "variable";
  const monthlyAmounts = record?.monthlyBillingAmounts || {};
  // If an override is explicitly supplied, use only that value.
  // This is important for Common pricing: an empty override means the
  // site allocation is not complete yet, so we must NOT fall back to the
  // full Common Subscription Fee stored on the record.
  const hasFixedMonthlyFeeOverride = fixedMonthlyFeeOverride !== undefined;
  const fixedMonthlyFee = hasFixedMonthlyFeeOverride
    ? parseAmountValue(fixedMonthlyFeeOverride)
    : parseAmountValue(record?.subscriptionFee);
  const lastVisibleMonthStart = getLatestNormalBillingMonthStart(asOfDate);
  if (!lastVisibleMonthStart) {
    return [];
  }

  let cursor = new Date(
    billingStartDate.getFullYear(),
    billingStartDate.getMonth(),
    1,
  );

  const rows = [];

  while (cursor <= lastVisibleMonthStart) {
    const year = cursor.getFullYear();
    const month = cursor.getMonth();
    const monthStart = new Date(year, month, 1);
    const monthEnd = new Date(year, month + 1, 0);

    const isFirstBillingMonth =
      year === billingStartDate.getFullYear() &&
      month === billingStartDate.getMonth();

    // Both monthly subscription types use the actual start date for the
    // first period; later periods begin on the first day of their month.
    const periodFrom = isFirstBillingMonth ? billingStartDate : monthStart;

    const periodTo = monthEnd;
    const monthKey = getBillingMonthKey(cursor);

    let amount = "";

    if (isVariable) {
      amount = normalizeValue(monthlyAmounts[monthKey]);
    } else if (fixedMonthlyFee !== null) {
      if (isFirstBillingMonth && billingStartDate.getDate() > 1) {
        const billableDays =
          monthEnd.getDate() - billingStartDate.getDate() + 1;

        amount = formatAmountValue(
          (fixedMonthlyFee / monthEnd.getDate()) * billableDays,
        );
      } else {
        amount = formatAmountValue(fixedMonthlyFee);
      }
    }

    const invoiceEligible = isNormalBillingMonthAvailable(cursor, asOfDate);

    rows.push({
      monthKey,
      billingMonth: getBillingMonthLabel(cursor),
      periodFrom: formatLocalDateValue(periodFrom),
      periodTo: formatLocalDateValue(periodTo),
      subscriptionType: normalizeValue(record?.subscriptionType),
      amount,
      invoiceEligible,
      invoiceEligibilityLabel: invoiceEligible
        ? "Eligible"
        : "Next Month",
    });

    cursor = new Date(year, month + 1, 1);
  }

  return rows;
}

function getCompletedMonthlyPeriodKeys(record, billingRecords = []) {
  const sourceRecordId = normalizeValue(record?.recordId);
  const screenCode = normalizeValue(record?.screenCode);
  const completedKeys = new Set();

  billingRecords.forEach((billingRecord) => {
    const matchesRecord =
      (sourceRecordId &&
        normalizeValue(billingRecord?.sourceRecordId) === sourceRecordId) ||
      (screenCode &&
        normalizeValue(billingRecord?.screenCode || billingRecord?.siteScope) ===
          screenCode);

    if (!matchesRecord) return;

    const status = normalizeValue(
      billingRecord?.submissionStatus || billingRecord?.billingVerificationStatus,
    ).toLowerCase();
    const isCompleted =
      Boolean(
        normalizeValue(billingRecord?.invoiceNumber) ||
          normalizeValue(billingRecord?.invoiceDate),
      ) ||
      status.includes("submitted") ||
      status.includes("completed") ||
      status.includes("billed");

    if (!isCompleted) return;

    const periodEntries = Array.isArray(billingRecord?.invoiceEntries)
      ? billingRecord.invoiceEntries
      : [billingRecord];

    periodEntries.forEach((entry) => {
      const periodFrom = parseLocalDateValue(
        entry?.periodFrom || billingRecord?.billingPeriodFrom,
      );
      const periodTo = parseLocalDateValue(
        entry?.periodTo || billingRecord?.billingPeriodTo || entry?.periodFrom,
      );

      if (!periodFrom || !periodTo || periodTo < periodFrom) return;

      let cursor = new Date(periodFrom.getFullYear(), periodFrom.getMonth(), 1);
      const finalMonth = new Date(periodTo.getFullYear(), periodTo.getMonth(), 1);

      while (cursor <= finalMonth) {
        completedKeys.add(getBillingMonthKey(cursor));
        cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
      }
    });
  });

  return completedKeys;
}

function getFirstTimeWaiverValidationError(draft, rows, completedPeriodKeys, billingStartDate) {
  if (draft.applicable !== "Yes") return "";

  const waiverFrom = parseLocalDateValue(draft.from);
  const waiverTo = parseLocalDateValue(draft.to);
  const reason = normalizeValue(draft.reason);
  const billingStart = parseLocalDateValue(billingStartDate);

  if (!waiverFrom) return "Waiver From is required.";
  if (!waiverTo) return "Waiver To is required.";
  if (!reason) return "Waiver Reason is required.";
  if (waiverTo < waiverFrom) return "Waiver To cannot be earlier than Waiver From.";
  if (billingStart && waiverFrom < billingStart) {
    return "Waiver From cannot be before the Billing Start Date.";
  }

  const firstMonth = new Date(waiverFrom.getFullYear(), waiverFrom.getMonth(), 1);
  const lastMonth = new Date(waiverTo.getFullYear(), waiverTo.getMonth(), 1);
  for (let cursor = firstMonth; cursor <= lastMonth; cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1)) {
    if (completedPeriodKeys.has(getBillingMonthKey(cursor))) {
      return "Waiver period cannot include an already billed period.";
    }
  }

  const overlapsPendingPeriod = rows.some((row) => {
    const periodFrom = parseLocalDateValue(row.periodFrom);
    const periodTo = parseLocalDateValue(row.periodTo);
    return periodFrom && periodTo && waiverFrom <= periodTo && waiverTo >= periodFrom;
  });

  return overlapsPendingPeriod ? "" : "Waiver period must overlap an existing pending billing period.";
}

function buildFirstTimePeriodBreakdown(rows, waiverDraft) {
  const waiverFrom = waiverDraft.applicable === "Yes" ? parseLocalDateValue(waiverDraft.from) : null;
  const waiverTo = waiverDraft.applicable === "Yes" ? parseLocalDateValue(waiverDraft.to) : null;

  return rows.map((row) => {
    const periodFrom = parseLocalDateValue(row.periodFrom);
    const periodTo = parseLocalDateValue(row.periodTo);
    const waived = Boolean(
      waiverFrom && waiverTo && periodFrom && periodTo && waiverFrom <= periodTo && waiverTo >= periodFrom,
    );
    const normalAmount = parseAmountValue(row.amount) || 0;

    return {
      ...row,
      treatment: waived ? "Waived" : "Billable",
      normalAmount: formatAmountValue(normalAmount),
      billableAmount: waived ? "0" : row.amount,
      invoiceEligibilityLabel: waived ? "Not Billable - Waiver" : row.invoiceEligibilityLabel,
    };
  });
}

function buildSubmittedFirstTimePeriodRows(record) {
  const breakdown = Array.isArray(record?.firstTimePeriodBreakdown)
    ? record.firstTimePeriodBreakdown
    : [];
  const billingCycleType = normalizeValue(record?.billingCycleType)
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (
    !breakdown.length ||
    (billingCycleType !== "first time" && !record?.firstTimeBillingSummary)
  ) {
    return [];
  }

  const savedEntries = Array.isArray(record.invoiceEntries)
    ? record.invoiceEntries
    : [];
  const savedByMonth = new Map(
    savedEntries.map((entry) => [normalizeValue(entry.entryId), entry]),
  );

  return breakdown
    .map((period, index) => {
      const periodFrom = normalizeValue(period?.periodFrom);
      const periodTo = normalizeValue(period?.periodTo);
      const monthKey =
        normalizeValue(period?.monthKey) ||
        (parseLocalDateValue(periodFrom)
          ? getBillingMonthKey(parseLocalDateValue(periodFrom))
          : `first-time-${index + 1}`);
      const existing = savedByMonth.get(monthKey) || {};
      const treatment = normalizeValue(period?.treatment);
      const isWaived = treatment.toLowerCase() === "waived";
      const billableAmount = normalizeValue(period?.billableAmount);

      return {
        ...existing,
        entryId: monthKey,
        billingMonth:
          normalizeValue(period?.billingMonth) ||
          formatBillingPeriod(periodFrom, periodTo),
        periodFrom,
        periodTo,
        invoiceNumber: normalizeValue(existing.invoiceNumber),
        invoiceDate: normalizeValue(existing.invoiceDate),
        invoiceAmountBeforeGST:
          billableAmount || normalizeValue(period?.amount),
        invoiceStatus:
          normalizeValue(existing.invoiceStatus) || "Pending Invoice",
        paymentStatus: isWaived
          ? "Waived"
          : normalizeValue(existing.paymentStatus) || "Pending",
        paymentReceivedDate: normalizeValue(existing.paymentReceivedDate),
        paymentRemarks: normalizeValue(existing.paymentRemarks),
        receiptNumber: normalizeValue(existing.receiptNumber),
        receiptDate: normalizeValue(existing.receiptDate),
        receivedAmount: normalizeValue(existing.receivedAmount),
        treatment,
        billableAmount,
        invoiceEligibilityLabel: normalizeValue(
          period?.invoiceEligibilityLabel,
        ),
        submittedPeriodTreatment: isWaived ? "Waived" : "Billable",
        submittedFirstTimePeriod: true,
      };
    })
    .filter((period) => period.periodFrom || period.periodTo);
}

function getCommonComplexBillingGroupKey(record) {
  return (
    normalizeValue(record?.stage1GroupId) ||
    normalizeValue(record?.complexCode) ||
    normalizeValue(record?.billingCode)
  );
}

function getCommonComplexBillingSiteKey(screenCode) {
  return normalizeValue(screenCode).toUpperCase();
}

function getSavedSiteBillingAllocation(record, commonComplexAllocations = {}) {
  if (!record) {
    return "";
  }

  const groupKey = getCommonComplexBillingGroupKey(record);
  const siteKey = getCommonComplexBillingSiteKey(record.screenCode);

  if (!groupKey || !siteKey) {
    return "";
  }

  return normalizeValue(
    commonComplexAllocations?.[groupKey]?.[siteKey]?.allocatedFee,
  );
}

function parseAmountValue(value) {
  const normalizedValue = normalizeValue(value).replace(/,/g, "");

  if (!normalizedValue) {
    return null;
  }

  if (!/^\d+(\.\d{1,2})?$/.test(normalizedValue)) {
    return null;
  }

  return Number(normalizedValue);
}

function formatAmountValue(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  const numericValue = Number(value);

  if (!Number.isFinite(numericValue)) {
    return "";
  }

  const roundedValue = Math.round(numericValue * 100) / 100;

  return Number.isInteger(roundedValue)
    ? String(roundedValue)
    : roundedValue.toFixed(2);
}

function formatInrAmount(value) {
  const amount = parseAmountValue(value);

  if (amount === null) {
    return "-";
  }

  return `\u20B9${amount.toLocaleString("en-IN", {
    minimumFractionDigits: Number.isInteger(amount) ? 0 : 2,
    maximumFractionDigits: 2,
  })}`;
}

function getCurrentLocalTimestamp() {
  return Date.now();
}

function getDownloadDateTimeStamp() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");

  return `${year}-${month}-${day}-${hours}${minutes}${seconds}`;
}

function formatTimestampValue(timestamp) {
  if (timestamp === null || timestamp === undefined || timestamp === "") {
    return "-";
  }

  const date = new Date(timestamp);

  if (Number.isNaN(date.getTime())) {
    return "-";
  }

  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  let hours = date.getHours();
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const suffix = hours >= 12 ? "PM" : "AM";

  hours = hours % 12 || 12;

  return `${day}-${month}-${year} ${String(hours).padStart(2, "0")}:${minutes} ${suffix}`;
}

function getApplicableBillingCycleBounds(record) {
  const billingStartDate = parseLocalDateValue(record?.billingStartDate);

  if (!billingStartDate) {
    return null;
  }

  const normalizedMode = normalizeValue(record?.subscriptionMode)
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const billingYear = billingStartDate.getFullYear();
  const billingMonth = billingStartDate.getMonth();

  if (
    normalizedMode === "yearly" ||
    normalizedMode === "annual" ||
    normalizedMode === "one year"
  ) {
    const cycleStart = new Date(
      billingStartDate.getFullYear(),
      billingStartDate.getMonth(),
      billingStartDate.getDate(),
    );
    const cycleEnd = addCalendarYears(cycleStart, 1);
    cycleEnd.setDate(cycleEnd.getDate() - 1);

    return {
      cycleStart,
      cycleEnd,
    };
  }

  if (
    normalizedMode === "half year" ||
    normalizedMode === "half yearly" ||
    normalizedMode === "semi annual"
  ) {
    const cycleStart = new Date(
      billingStartDate.getFullYear(),
      billingStartDate.getMonth(),
      billingStartDate.getDate(),
    );
    const cycleEnd = addCalendarMonths(cycleStart, 6);
    cycleEnd.setDate(cycleEnd.getDate() - 1);

    return {
      cycleStart,
      cycleEnd,
    };
  }

  return {
    cycleStart: new Date(billingYear, billingMonth, 1),
    cycleEnd: new Date(billingYear, billingMonth + 1, 0),
  };
}

function getInclusiveDateDifference(startDate, endDate) {
  if (!startDate || !endDate) {
    return null;
  }

  const differenceInMs = endDate.getTime() - startDate.getTime();

  if (differenceInMs < 0) {
    return null;
  }

  return Math.floor(differenceInMs / (1000 * 60 * 60 * 24)) + 1;
}


function getBillingResumeDate(record) {
  const directActiveFrom = parseLocalDateValue(record?.billingActiveFromDate);

  if (directActiveFrom) {
    return directActiveFrom;
  }

  const history = Array.isArray(record?.closureHistory)
    ? record.closureHistory
    : [];

  const latestRestore = history
    .filter(
      (entry) =>
        normalizeValue(entry?.newStatus).toLowerCase() === "active" &&
        parseLocalDateValue(entry?.effectiveDate),
    )
    .sort((left, right) => {
      const leftDate = parseLocalDateValue(left?.effectiveDate);
      const rightDate = parseLocalDateValue(right?.effectiveDate);
      return (rightDate?.getTime() || 0) - (leftDate?.getTime() || 0);
    })[0];

  return latestRestore
    ? parseLocalDateValue(latestRestore.effectiveDate)
    : null;
}



function getSortedClosureHistory(record) {
  const history = Array.isArray(record?.closureHistory)
    ? record.closureHistory
    : [];

  return [...history]
    .map((entry, index) => ({
      ...entry,
      __historyIndex: index,
    }))
    .sort((left, right) => {
      const leftUpdated =
        parseLocalDateValue(left?.updatedAt) ||
        parseLocalDateValue(left?.effectiveDate);
      const rightUpdated =
        parseLocalDateValue(right?.updatedAt) ||
        parseLocalDateValue(right?.effectiveDate);

      const leftTime = leftUpdated?.getTime() || 0;
      const rightTime = rightUpdated?.getTime() || 0;

      if (rightTime !== leftTime) {
        return rightTime - leftTime;
      }

      return right.__historyIndex - left.__historyIndex;
    });
}

function getBillingPauseDate(record) {
  const directPauseFrom = parseLocalDateValue(
    record?.pauseFromDate || record?.closureEffectiveDate,
  );

  if (directPauseFrom) {
    return directPauseFrom;
  }

  const history = Array.isArray(record?.closureHistory)
    ? record.closureHistory
    : [];

  const latestPause = history
    .filter((entry) => {
      const nextStatus = normalizeValue(entry?.newStatus).toLowerCase();
      return (
        [
          "billing paused",
          "site inactive",
          "site removed",
          "contract closed",
          "inactive",
        ].includes(nextStatus) &&
        parseLocalDateValue(entry?.effectiveDate)
      );
    })
    .sort((left, right) => {
      const leftDate = parseLocalDateValue(left?.effectiveDate);
      const rightDate = parseLocalDateValue(right?.effectiveDate);
      return (rightDate?.getTime() || 0) - (leftDate?.getTime() || 0);
    })[0];

  return latestPause
    ? parseLocalDateValue(latestPause.effectiveDate)
    : null;
}

function isFinalLifecycleBillingPeriod(
  record,
  periodFromValue,
  periodToValue,
  asOfDate = new Date(),
) {
  const periodFrom = parseLocalDateValue(periodFromValue);
  const periodTo = parseLocalDateValue(periodToValue);
  const asOf = getBillingAsOfDate(asOfDate);
  const effectiveEnd = getBillingPauseDate(record);
  const status = normalizeValue(
    record?.billingLifecycleReason ||
      record?.billingStatusReason ||
      record?.currentBillingStatus ||
      record?.billingLifecycleStatus,
  ).toLowerCase();
  const isFinalStatus = [
    "quit",
    "cancelled",
    "canceled",
    "site inactive",
    "site removed",
    "contract closed",
    "inactive",
    "closed",
  ].includes(status);

  return Boolean(
    isFinalStatus &&
      periodFrom &&
      periodTo &&
      effectiveEnd &&
      asOf &&
      effectiveEnd <= asOf &&
      effectiveEnd >= periodFrom &&
      effectiveEnd <=
        new Date(periodTo.getFullYear(), periodTo.getMonth() + 1, 1),
  );
}

function getLifecycleAdjustedMonthlyPeriod(
  record,
  monthStartValue,
  monthEndValue,
) {
  const monthStart = parseLocalDateValue(monthStartValue);
  const monthEnd = parseLocalDateValue(monthEndValue);

  if (!monthStart || !monthEnd || monthEnd < monthStart) {
    return {
      periodFrom: monthStartValue || "",
      periodTo: monthEndValue || "",
      billable: Boolean(monthStart && monthEnd),
    };
  }

  const history = Array.isArray(record?.closureHistory)
    ? record.closureHistory
    : [];

  const lifecycleEvents = history
    .map((entry) => ({
      ...entry,
      eventDate: parseLocalDateValue(entry?.effectiveDate),
      normalizedStatus: normalizeValue(entry?.newStatus).toLowerCase(),
    }))
    .filter((entry) => entry.eventDate)
    .sort((left, right) => left.eventDate - right.eventDate);

  // Backward compatible fallback for legacy records.
  const directPauseDate = parseLocalDateValue(
    record?.pauseFromDate || record?.closureEffectiveDate,
  );
  const directResumeDate = parseLocalDateValue(record?.billingActiveFromDate);

  if (
    directPauseDate &&
    !lifecycleEvents.some(
      (event) =>
        event.eventDate.getTime() === directPauseDate.getTime() &&
        [
          "billing paused",
          "site inactive",
          "site removed",
          "contract closed",
          "inactive",
        ].includes(event.normalizedStatus),
    )
  ) {
    lifecycleEvents.push({
      eventDate: directPauseDate,
      normalizedStatus: normalizeValue(
        record?.billingLifecycleReason || record?.billingStatusReason || "inactive",
      ).toLowerCase(),
    });
  }

  if (
    directResumeDate &&
    !lifecycleEvents.some(
      (event) =>
        event.eventDate.getTime() === directResumeDate.getTime() &&
        event.normalizedStatus === "active",
    )
  ) {
    lifecycleEvents.push({
      eventDate: directResumeDate,
      normalizedStatus: "active",
    });
  }

  lifecycleEvents.sort((left, right) => left.eventDate - right.eventDate);

  const isInactiveStatus = (status) =>
    [
      "billing paused",
      "site inactive",
      "site removed",
      "contract closed",
      "inactive",
    ].includes(status);

  // Determine status at the beginning of this month.
  let isActive = true;

  lifecycleEvents.forEach((event) => {
    if (event.eventDate < monthStart) {
      if (event.normalizedStatus === "active") {
        isActive = true;
      } else if (isInactiveStatus(event.normalizedStatus)) {
        isActive = false;
      }
    }
  });

  const billableSegments = [];
  let segmentStart = isActive ? new Date(monthStart) : null;

  lifecycleEvents.forEach((event) => {
    if (event.eventDate < monthStart || event.eventDate > monthEnd) {
      return;
    }

    if (isInactiveStatus(event.normalizedStatus)) {
      if (segmentStart) {
        const dayBeforeInactive = new Date(event.eventDate);
        dayBeforeInactive.setDate(dayBeforeInactive.getDate() - 1);

        if (dayBeforeInactive >= segmentStart) {
          billableSegments.push({
            from: new Date(segmentStart),
            to: dayBeforeInactive,
          });
        }
      }

      segmentStart = null;
      isActive = false;
      return;
    }

    if (event.normalizedStatus === "active") {
      if (!segmentStart) {
        segmentStart = new Date(event.eventDate);
      }
      isActive = true;
    }
  });

  if (segmentStart && segmentStart <= monthEnd) {
    billableSegments.push({
      from: new Date(segmentStart),
      to: new Date(monthEnd),
    });
  }

  if (billableSegments.length === 0) {
    const inactiveIntervals = [];

    lifecycleEvents.forEach((event, index) => {
      if (!isInactiveStatus(event.normalizedStatus)) return;

      const nextActiveEvent = lifecycleEvents
        .slice(index + 1)
        .find((candidate) => candidate.normalizedStatus === "active");
      const intervalStart = event.eventDate;
      const intervalEnd = nextActiveEvent?.eventDate || null;
      const overlapsMonth =
        intervalStart <= monthEnd && (!intervalEnd || intervalEnd > monthStart);

      if (overlapsMonth) inactiveIntervals.push({ from: intervalStart, to: intervalEnd });
    });

    const relevantInactiveInterval = inactiveIntervals[inactiveIntervals.length - 1] || null;
    const lifecycleRemarks = relevantInactiveInterval
      ? `Site paused/inactive from ${formatBillingDateDisplay(formatLocalDateValue(relevantInactiveInterval.from))}${
          relevantInactiveInterval.to
            ? ` and restored to Active from ${formatBillingDateDisplay(formatLocalDateValue(relevantInactiveInterval.to))}.`
            : "."
        }`
      : "Site is paused/inactive for this billing month.";

    return {
      periodFrom: "",
      periodTo: "",
      billable: false,
      lifecycleStatus: "Paused / Not Billable",
      lifecycleRemarks,
      lifecyclePeriodLabel: `${formatBillingDateDisplay(formatLocalDateValue(monthStart))} to ${formatBillingDateDisplay(formatLocalDateValue(monthEnd))}`,
    };
  }

  // Current UI supports one monthly billing period per row. For the expected
  // deactivate/reactivate workflow there will be a single billable segment
  // in any given month.
  const segment = billableSegments[0];

  return {
    periodFrom: formatLocalDateValue(segment.from),
    periodTo: formatLocalDateValue(segment.to),
    billable: true,
    lifecycleStatus: "Billable",
    lifecycleRemarks: "",
    lifecyclePeriodLabel: "",
  };
}

function getEffectiveBillingPeriodFrom(record, requestedPeriodFrom) {
  const requestedFrom = parseLocalDateValue(requestedPeriodFrom);

  if (!requestedFrom) {
    return requestedPeriodFrom || "";
  }

  const requestedMonthEnd = new Date(
    requestedFrom.getFullYear(),
    requestedFrom.getMonth() + 1,
    0,
  );

  const adjusted = getLifecycleAdjustedMonthlyPeriod(
    record,
    formatLocalDateValue(requestedFrom),
    formatLocalDateValue(requestedMonthEnd),
  );

  return adjusted.billable ? adjusted.periodFrom : "";
}

function getProratedFixedMonthlyAmount(amount, periodFrom, periodTo) {
  const numericAmount = parseAmountValue(amount);
  const fromDate = parseLocalDateValue(periodFrom);
  const toDate = parseLocalDateValue(periodTo);

  if (numericAmount === null || !fromDate || !toDate || toDate < fromDate) {
    return "";
  }

  // Only prorate within one calendar month. Recurring monthly rows are
  // generated month-wise, so this is the expected path.
  if (
    fromDate.getFullYear() !== toDate.getFullYear() ||
    fromDate.getMonth() !== toDate.getMonth()
  ) {
    return formatAmountValue(numericAmount);
  }

  const monthEnd = new Date(
    fromDate.getFullYear(),
    fromDate.getMonth() + 1,
    0,
  );
  const activeDays = getInclusiveDateDifference(fromDate, toDate);

  if (!activeDays) {
    return "";
  }

  return formatAmountValue(
    (numericAmount / monthEnd.getDate()) * activeDays,
  );
}

function getStandaloneCurrentBillingCharge(record) {
  const subscriptionFee = parseAmountValue(record?.subscriptionFee);
  const billingPeriodFrom = parseLocalDateValue(record?.billingPeriodFrom);
  const billingPeriodTo = parseLocalDateValue(record?.billingPeriodTo);
  const billingStartDate = parseLocalDateValue(record?.billingStartDate);

  if (
    subscriptionFee === null ||
    !billingPeriodFrom ||
    !billingPeriodTo ||
    !billingStartDate
  ) {
    return "";
  }

  const normalizedMode = normalizeValue(record?.subscriptionMode)
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  // For non-monthly plans, keep existing subscription fee behavior.
  if (
    normalizedMode !== "" &&
    normalizedMode !== "monthly" &&
    normalizedMode !== "month"
  ) {
    return formatAmountValue(subscriptionFee);
  }

  let total = 0;
  let cursor = new Date(
    billingPeriodFrom.getFullYear(),
    billingPeriodFrom.getMonth(),
    1,
  );
  const endMonth = new Date(
    billingPeriodTo.getFullYear(),
    billingPeriodTo.getMonth(),
    1,
  );

  while (cursor <= endMonth) {
    const year = cursor.getFullYear();
    const month = cursor.getMonth();
    const monthStart = new Date(year, month, 1);
    const monthEnd = new Date(year, month + 1, 0);

    const effectiveStart =
      year === billingPeriodFrom.getFullYear() &&
      month === billingPeriodFrom.getMonth()
        ? billingPeriodFrom
        : monthStart;

    const effectiveEnd =
      year === billingPeriodTo.getFullYear() &&
      month === billingPeriodTo.getMonth()
        ? billingPeriodTo
        : monthEnd;

    const billableDays = getInclusiveDateDifference(
      effectiveStart,
      effectiveEnd,
    );
    const totalMonthDays = monthEnd.getDate();

    if (billableDays && totalMonthDays) {
      total += (subscriptionFee / totalMonthDays) * billableDays;
    }

    cursor = new Date(year, month + 1, 1);
  }

  return formatAmountValue(total);
}

function getCommonComplexCurrentBillingCharge(record, allocatedFee) {
  const allocatedAmount = parseAmountValue(allocatedFee);
  const cycleBounds = getApplicableBillingCycleBounds(record);
  const billableStartDate = parseLocalDateValue(record?.billingStartDate);
  const billableEndDate = parseLocalDateValue(record?.billingPeriodTo);

  if (allocatedAmount === null || !cycleBounds || !billableStartDate) {
    return "";
  }

  const effectiveBillableEndDate =
    billableEndDate && billableEndDate < cycleBounds.cycleEnd
      ? billableEndDate
      : cycleBounds.cycleEnd;
  const billableDays = getInclusiveDateDifference(
    billableStartDate,
    effectiveBillableEndDate,
  );
  const totalCycleDays = getInclusiveDateDifference(
    cycleBounds.cycleStart,
    cycleBounds.cycleEnd,
  );

  if (!billableDays || !totalCycleDays) {
    return "";
  }

  const proratedAmount = (allocatedAmount / totalCycleDays) * billableDays;
  return formatAmountValue(proratedAmount);
}

function isVariableSubscriptionRecord(record) {
  return normalizeValue(record?.subscriptionType).toLowerCase() === "variable";
}

function isCommonComplexBillingRecord(record) {
  const pricingMethod = normalizeValue(record?.pricingMethod).toLowerCase();
  const complexCode = normalizeValue(record?.complexCode);
  const stage1GroupRows = Array.isArray(record?.stage1GroupRows)
    ? record.stage1GroupRows
    : [];

  // The currently saved pricingMethod is authoritative. For Common pricing,
  // the Stage 1 Subscription Fee is the total fee for the selected complex
  // screens and Stage 3 must allocate that total site-wise before validation.
  // Historical pricingGroups/allocationRows can remain on a corrected record;
  // they must not cause a true Common record to bypass the allocation step.
  return (
    pricingMethod === "common" &&
    (Boolean(complexCode) || stage1GroupRows.length > 1)
  );
}

function isCommonComplexSiteCurrentlyBillable(record) {
  if (!record || isFoCOriginRecord(record)) {
    return false;
  }

  const billingStartDate = parseLocalDateValue(record.billingStartDate);
  if (!billingStartDate) {
    return false;
  }

  const today = new Date();
  const todayStart = new Date(
    today.getFullYear(),
    today.getMonth(),
    today.getDate(),
  );
  const lifecycleStatus = normalizeValue(record.billingLifecycleStatus)
    .toLowerCase();
  const lifecycleReason = normalizeValue(record.billingLifecycleReason);

  return (
    billingStartDate <= todayStart &&
    lifecycleStatus !== "inactive" &&
    lifecycleStatus !== "paused" &&
    !lifecycleReason
  );
}

function getCommonComplexAllSiteRows(
  record,
  stage3Records = [],
  commonComplexAllocations = {},
  additionalMembershipRecords = [],
) {
  if (!isCommonComplexBillingRecord(record)) {
    return [];
  }

  const groupKey = getCommonComplexBillingGroupKey(record);
  const relatedStage3Records = stage3Records.filter(
    (candidate) => getCommonComplexBillingGroupKey(candidate) === groupKey,
  );
  const relatedMembershipRecords = additionalMembershipRecords.filter(
    (candidate) => getCommonComplexBillingGroupKey(candidate) === groupKey,
  );
  const sourceRows = [
    ...(Array.isArray(record?.stage1GroupRows) ? record.stage1GroupRows : []),
    ...relatedStage3Records.flatMap((candidate) =>
      Array.isArray(candidate?.stage1GroupRows) ? candidate.stage1GroupRows : [],
    ),
    ...relatedMembershipRecords,
    ...relatedStage3Records,
  ];
  const stage3RecordBySiteKey = new Map(
    relatedStage3Records.map((candidate) => [
      getCommonComplexBillingSiteKey(candidate.screenCode),
      candidate,
    ]),
  );
  const allocationGroup = commonComplexAllocations[groupKey] || {};
  const seenSiteKeys = new Set();

  return sourceRows
    .map((sourceRow) => {
      const siteKey = getCommonComplexBillingSiteKey(sourceRow?.screenCode);
      if (!siteKey || seenSiteKeys.has(siteKey)) {
        return null;
      }

      seenSiteKeys.add(siteKey);
      const relatedRecord = stage3RecordBySiteKey.get(siteKey) || null;
      const subscriptionApplicable = relatedRecord
        ? isSubscriptionApplicableRecord(relatedRecord)
        : !isFoCOriginRecord(sourceRow) && isSubscriptionApplicableRecord(record);
      const isFoC =
        isFoCOriginRecord(sourceRow) ||
        isFoCOriginRecord(relatedRecord) ||
        normalizeValue(sourceRow?.commercialApplicable).toLowerCase() ===
          "false" ||
        normalizeValue(relatedRecord?.commercialApplicable).toLowerCase() ===
          "false";
      const isBillable =
        !isFoC && isCommonComplexSiteCurrentlyBillable(relatedRecord);
      const lifecycleStatus = normalizeValue(
        relatedRecord?.billingLifecycleStatus,
      ).toLowerCase();
      const allocatedFee = isFoC
        ? "0"
        : normalizeValue(allocationGroup[siteKey]?.allocatedFee);

      return {
        siteKey,
        screenCode: normalizeValue(sourceRow?.screenCode),
        screenName:
          normalizeValue(sourceRow?.screenName) ||
          normalizeValue(relatedRecord?.screenName),
        billingStartDate: normalizeValue(relatedRecord?.billingStartDate),
        allocatedFee,
        isFoC,
        isBillable,
        status: isFoC
          ? "FoC"
          : isBillable
            ? "Active"
            : lifecycleStatus === "inactive" || lifecycleStatus === "paused"
              ? "Inactive"
              : "Pending",
        relatedRecord,
        subscriptionApplicable,
      };
    })
    .filter(Boolean);
}

function getCommonComplexPricingSummary(
  record,
  stage3Records = [],
  commonComplexAllocations = {},
  additionalMembershipRecords = [],
) {
  if (!isCommonComplexBillingRecord(record)) {
    return null;
  }

  const screenRows = getCommonComplexAllSiteRows(
    record,
    stage3Records,
    commonComplexAllocations,
    additionalMembershipRecords,
  );
  const sumRows = (rows) =>
    rows.reduce((total, row) => {
      const fee = parseAmountValue(row.allocatedFee);
      return total + (fee === null ? 0 : fee);
    }, 0);

  const activeRows = screenRows.filter((row) => row.isBillable);
  const pendingRows = screenRows.filter((row) => row.status === "Pending");

  return {
    groupKey: getCommonComplexBillingGroupKey(record),
    billingCode: normalizeValue(record.billingCode),
    complexCode: normalizeValue(record.complexCode),
    screenRows,
    activeBillableScreens: activeRows.length,
    currentCombinedBillingAmount: formatAmountValue(sumRows(activeRows)),
    configuredCommonTotal: normalizeValue(
      resolveCommonComplexSubscriptionFee(
        record,
        stage3Records,
        commonComplexAllocations,
      ),
    ),
    pendingBillableAmount: formatAmountValue(sumRows(pendingRows)),
    screensAwaitingBilling: pendingRows.length,
    focScreens: screenRows.filter((row) => row.isFoC).length,
  };
}

function getCommonComplexBillingRows(
  record,
  stage3Records = [],
  commonComplexAllocations = {},
  additionalMembershipRecords = [],
) {
  if (!isCommonComplexBillingRecord(record)) {
    return [];
  }

  const groupKey = getCommonComplexBillingGroupKey(record);
  const stage1GroupRows = Array.isArray(record?.stage1GroupRows)
    ? record.stage1GroupRows
    : [];
  const relatedStage3Records = stage3Records.filter(
    (candidate) => getCommonComplexBillingGroupKey(candidate) === groupKey,
  );
  const relatedMembershipRecords = additionalMembershipRecords.filter(
    (candidate) => getCommonComplexBillingGroupKey(candidate) === groupKey,
  );
  const billingGroupRows = [
    ...stage1GroupRows,
    ...relatedStage3Records.flatMap((candidate) =>
      Array.isArray(candidate?.stage1GroupRows) ? candidate.stage1GroupRows : [],
    ),
    ...relatedMembershipRecords,
    ...relatedStage3Records,
  ];
  const stage3RecordBySiteKey = new Map(
    [...relatedMembershipRecords, ...relatedStage3Records].map((candidate) => [
      getCommonComplexBillingSiteKey(candidate.screenCode),
      candidate,
    ]),
  );
  const allocationGroup = commonComplexAllocations[groupKey] || {};
  const seenSiteKeys = new Set();

  return billingGroupRows
    .map((row) => {
      const siteKey = getCommonComplexBillingSiteKey(row?.screenCode);

      if (!siteKey || seenSiteKeys.has(siteKey)) {
        return null;
      }

      seenSiteKeys.add(siteKey);

      const relatedRecord = stage3RecordBySiteKey.get(siteKey) || null;
      const subscriptionApplicable = relatedRecord
        ? isSubscriptionApplicableRecord(relatedRecord)
        : !isFoCOriginRecord(row) && isSubscriptionApplicableRecord(record);

      // FoC screens are outside the commercial subscription allocation.
      // Stage 1 common pricing must be split only across billable/non-FoC
      // screens in the selected complex. Check both the Stage 1 snapshot row
      // and its live Stage 3 record because either can carry the FoC marker.
      if (
        !subscriptionApplicable ||
        isFoCOriginRecord(row) ||
        isFoCOriginRecord(relatedRecord) ||
        normalizeValue(row?.commercialApplicable).toLowerCase() === "false" ||
        normalizeValue(relatedRecord?.commercialApplicable).toLowerCase() === "false"
      ) {
        return null;
      }

      const allocationEntry = allocationGroup[siteKey] || {};
      const allocatedFee = normalizeValue(allocationEntry.allocatedFee);
      const billingVerificationStatus = normalizeValue(
        relatedRecord?.billingVerificationStatus,
      );
      const firstBillingCompleted =
        billingVerificationStatus === "Submitted to Billing Team" ||
        billingVerificationStatus === "Sent to Billing Team" ||
        Boolean(relatedRecord?.billingVerifiedAt);

      return {
        siteKey,
        screenCode: normalizeValue(row?.screenCode),
        screenName:
          normalizeValue(row?.screenName) || normalizeValue(relatedRecord?.screenName),
        billingStartDate: normalizeValue(relatedRecord?.billingStartDate),
        billingStatus: relatedRecord ? "Ready" : "Pending",
        firstBillingCompleted,
        allocatedFee,
        locked: Boolean(allocationEntry.locked),
        billingAllocationCreatedAt: allocationEntry.billingAllocationCreatedAt || "",
        billingAllocationUpdatedAt: allocationEntry.billingAllocationUpdatedAt || "",
        currentBillingCharge: relatedRecord
          ? getCommonComplexCurrentBillingCharge(relatedRecord, allocatedFee)
          : "",
        isBillable: isCommonComplexSiteCurrentlyBillable(relatedRecord),
        relatedRecord,
      };
    })
    .filter(Boolean);
}


function getCommonComplexBillingAlignment(
  record,
  stage3Records = [],
  commonComplexAllocations = {},
  additionalMembershipRecords = [],
) {
  if (
    !record ||
    !isCommonComplexBillingRecord(record) ||
    isVariableSubscriptionRecord(record)
  ) {
    return {
      allApplicableSitesStarted: false,
      sameBillingStartDate: false,
      commonBillingStartDate: "",
      billingMode: "Site-wise Allocation",
    };
  }

  const rows = getCommonComplexBillingRows(
    record,
    stage3Records,
    commonComplexAllocations,
    additionalMembershipRecords,
  );

  if (rows.length === 0) {
    return {
      allApplicableSitesStarted: false,
      sameBillingStartDate: false,
      commonBillingStartDate: "",
      billingMode: "Site-wise Allocation",
    };
  }

  const allApplicableSitesStarted = rows.every(
    (row) => row.isBillable && parseLocalDateValue(row.billingStartDate),
  );
  const billingStartDates = rows
    .map((row) => normalizeValue(row.billingStartDate))
    .filter(Boolean);
  const uniqueBillingStartDates = new Set(billingStartDates);
  const sameBillingStartDate =
    allApplicableSitesStarted &&
    billingStartDates.length === rows.length &&
    uniqueBillingStartDates.size === 1;

  return {
    allApplicableSitesStarted,
    sameBillingStartDate,
    commonBillingStartDate: sameBillingStartDate ? billingStartDates[0] : "",
    billingMode: sameBillingStartDate
      ? "Combined Complex Billing"
      : "Site-wise Allocation",
  };
}

function canUseCombinedCommonBilling(
  record,
  stage3Records = [],
  commonComplexAllocations = {},
  additionalMembershipRecords = [],
) {
  if (!record || isVariableSubscriptionRecord(record)) {
    return false;
  }

  const alignment = getCommonComplexBillingAlignment(
    record,
    stage3Records,
    commonComplexAllocations,
    additionalMembershipRecords,
  );

  if (!alignment.sameBillingStartDate) {
    return false;
  }

  const billingRows = getCommonComplexBillingRows(
    record,
    stage3Records,
    commonComplexAllocations,
    additionalMembershipRecords,
  );

  if (billingRows.some((row) => !row.firstBillingCompleted)) {
    return false;
  }

  return (
    getCommonComplexBillingValidationError(
      record,
      stage3Records,
      commonComplexAllocations,
      additionalMembershipRecords,
    ) === ""
  );
}

function resolveCommonComplexSubscriptionFee(
  record
) {
  if (!record || !isCommonComplexBillingRecord(record)) {
    return normalizeValue(record?.subscriptionFee);
  }
  // Hydration supplies the latest active Common CommercialTerm. Allocation is
  // validated against that fee; it must never decide which fee is current.
  return normalizeValue(record?.subscriptionFee);
}

function getCommonComplexBillingValidationError(
  record,
  stage3Records = [],
  commonComplexAllocations = {},
  additionalMembershipRecords = [],
) {
  if (
    !isCommonComplexBillingRecord(record) ||
    isVariableSubscriptionRecord(record)
  ) {
    return "";
  }

  const commonSubscriptionFee = parseAmountValue(
    resolveCommonComplexSubscriptionFee(
      record,
      stage3Records,
      commonComplexAllocations,
      additionalMembershipRecords,
    ),
  );

  if (commonSubscriptionFee === null) {
    return "Common Subscription Fee is invalid.";
  }

  const billingRows = getCommonComplexBillingRows(
    record,
    stage3Records,
    commonComplexAllocations,
    additionalMembershipRecords,
  );

  let allocationTotal = 0;

  for (const row of billingRows) {
    const allocatedFee = parseAmountValue(row.allocatedFee);

    if (allocatedFee === null) {
      return `Allocation is required for ${row.screenCode}. All included screens must be allocated before saving.`;
    }

    allocationTotal += allocatedFee;
  }

  if (Math.abs(allocationTotal - commonSubscriptionFee) > 0.01) {
    return `Site allocation total ${formatAmountValue(
      allocationTotal,
    )} does not match the Common Subscription Fee ${formatAmountValue(
      commonSubscriptionFee,
    )}.`;
  }

  return "";
}

function resolveCommonComplexBillingStatus(
  record,
  stage3Records = [],
  commonComplexAllocations = {},
  additionalMembershipRecords = [],
) {
  if (!isCommonComplexBillingRecord(record)) {
    return "";
  }

  const billingRows = getCommonComplexBillingRows(
    record,
    stage3Records,
    commonComplexAllocations,
    additionalMembershipRecords,
  );

  const validationError = getCommonComplexBillingValidationError(
    record,
    stage3Records,
    commonComplexAllocations,
    additionalMembershipRecords,
  );

  if (validationError) {
    return "Allocation Pending";
  }

  const alignment = getCommonComplexBillingAlignment(
    record,
    stage3Records,
    commonComplexAllocations,
    additionalMembershipRecords,
  );

  if (!alignment.allApplicableSitesStarted) {
    return "Allocated Site-wise Billing";
  }

  if (!alignment.sameBillingStartDate) {
    return "Allocated Site-wise Billing";
  }

  if (billingRows.some((row) => row.isBillable && !row.firstBillingCompleted)) {
    return "Combined Billing Pending";
  }

  return "Combined Complex Billing";
}

function getLatestSubmittedBillingRecord(record, billingRecords = []) {
  const sourceRecordId = normalizeValue(record?.recordId);
  const screenCode = normalizeValue(record?.screenCode);

  return billingRecords
    .filter((billingRecord) => {
      const sameSource =
        sourceRecordId &&
        normalizeValue(billingRecord?.sourceRecordId) === sourceRecordId;
      const sameSiteFallback =
        !normalizeValue(billingRecord?.sourceRecordId) &&
        screenCode &&
        normalizeValue(billingRecord?.screenCode) === screenCode;
      const submitted = [
        "Submitted to Billing Team",
        "Sent to Billing Team",
      ].includes(normalizeValue(billingRecord?.submissionStatus));

      return submitted && (sameSource || sameSiteFallback);
    })
    .sort((left, right) => {
      const leftDate = parseLocalDateValue(left?.billingPeriodTo);
      const rightDate = parseLocalDateValue(right?.billingPeriodTo);
      return (rightDate?.getTime() || 0) - (leftDate?.getTime() || 0);
    })[0] || null;
}


function getPriceChangeRequest(record) {
  return record?.priceChangeRequest && typeof record.priceChangeRequest === "object"
    ? record.priceChangeRequest
    : null;
}

function isPriceChangeRequestOpen(record) {
  const request = getPriceChangeRequest(record);
  if (!request) return false;

  const status = normalizeValue(request.status).toLowerCase();
  return ["draft", "pending validation", "pending approval"].includes(status);
}

function doesPriceChangeBlockBillingPeriod(record, periodFrom, periodTo) {
  if (!isPriceChangeRequestOpen(record)) return false;

  const request = getPriceChangeRequest(record);
  const effectiveDate = parseLocalDateValue(request?.effectiveDate);
  const periodEnd = parseLocalDateValue(periodTo);

  // A newly initiated change without a valid effective date remains blocked
  // until the mandatory Effective Date is completed and validated.
  if (!effectiveDate || !periodEnd) return true;

  return effectiveDate <= periodEnd;
}

function getPriceChangeBillingBlock(
  record,
  stage3Records = [],
  periodFrom = "",
  periodTo = "",
) {
  if (!record) {
    return { blocked: false, scope: "", label: "" };
  }

  if (isCommonComplexBillingRecord(record)) {
    const groupKey = getCommonComplexBillingGroupKey(record);
    const relatedRecords = stage3Records.filter(
      (candidate) => getCommonComplexBillingGroupKey(candidate) === groupKey,
    );
    const blockedRecord = relatedRecords.find((candidate) =>
      doesPriceChangeBlockBillingPeriod(candidate, periodFrom, periodTo),
    );

    if (blockedRecord) {
      return {
        blocked: true,
        scope: "Common Complex",
        label:
          normalizeValue(record?.complexCode) ||
          normalizeValue(record?.billingName) ||
          normalizeValue(record?.billingCode) ||
          "This complex",
      };
    }
  }

  if (doesPriceChangeBlockBillingPeriod(record, periodFrom, periodTo)) {
    return {
      blocked: true,
      scope: "Site",
      label:
        normalizeValue(record?.screenName) ||
        normalizeValue(record?.screenCode) ||
        normalizeValue(record?.billingCode) ||
        "This site",
    };
  }

  return { blocked: false, scope: "", label: "" };
}

function getLatestEffectivePriceChange(record, periodTo) {
  const periodEnd = parseLocalDateValue(periodTo);
  if (!periodEnd) return null;

  const history = Array.isArray(record?.priceChangeHistory)
    ? record.priceChangeHistory
    : [];

  return history
    .filter((entry) => {
      const effectiveDate = parseLocalDateValue(entry?.effectiveDate);
      return effectiveDate && effectiveDate <= periodEnd;
    })
    .sort((left, right) => {
      const leftDate = parseLocalDateValue(left?.effectiveDate);
      const rightDate = parseLocalDateValue(right?.effectiveDate);
      return (rightDate?.getTime() || 0) - (leftDate?.getTime() || 0);
    })[0] || null;
}

function getPriceChangeAdjustedFixedAmount(
  record,
  periodFrom,
  periodTo,
  fallbackFee,
) {
  const fromDate = parseLocalDateValue(periodFrom);
  const toDate = parseLocalDateValue(periodTo);
  const fallbackAmount = parseAmountValue(fallbackFee);

  if (!fromDate || !toDate || toDate < fromDate || fallbackAmount === null) {
    return "";
  }

  const priceChange = getLatestEffectivePriceChange(record, periodTo);
  if (!priceChange) {
    const normalizedMode = normalizeValue(record?.subscriptionMode)
      .toLowerCase()
      .replace(/[-_]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    const isMonthly = ["", "monthly", "month"].includes(normalizedMode);
    return isMonthly
      ? getProratedFixedMonthlyAmount(fallbackFee, periodFrom, periodTo)
      : formatAmountValue(fallbackAmount);
  }

  const effectiveDate = parseLocalDateValue(priceChange?.effectiveDate);
  const previousFee = parseAmountValue(
    priceChange?.previousSiteFee ?? priceChange?.previousFee,
  );
  const newFee = parseAmountValue(
    priceChange?.newSiteFee ?? priceChange?.newFee ?? fallbackFee,
  );

  if (!effectiveDate || previousFee === null || newFee === null) {
    return formatAmountValue(fallbackAmount);
  }

  const normalizedMode = normalizeValue(record?.subscriptionMode)
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const isMonthly = ["", "monthly", "month"].includes(normalizedMode);

  let cycleStart;
  let cycleEnd;

  if (isMonthly) {
    cycleStart = new Date(fromDate.getFullYear(), fromDate.getMonth(), 1);
    cycleEnd = new Date(fromDate.getFullYear(), fromDate.getMonth() + 1, 0);
  } else {
    cycleStart = new Date(fromDate);
    cycleEnd = new Date(toDate);
  }

  const totalCycleDays = getInclusiveDateDifference(cycleStart, cycleEnd);
  if (!totalCycleDays) return "";

  const calculateSegment = (segmentFrom, segmentTo, fullCycleFee) => {
    if (!segmentFrom || !segmentTo || segmentTo < segmentFrom) return 0;
    const days = getInclusiveDateDifference(segmentFrom, segmentTo) || 0;
    return (fullCycleFee / totalCycleDays) * days;
  };

  if (effectiveDate <= fromDate) {
    return formatAmountValue(calculateSegment(fromDate, toDate, newFee));
  }

  if (effectiveDate > toDate) {
    return formatAmountValue(calculateSegment(fromDate, toDate, previousFee));
  }

  const oldPeriodEnd = new Date(effectiveDate);
  oldPeriodEnd.setDate(oldPeriodEnd.getDate() - 1);

  const oldAmount = calculateSegment(fromDate, oldPeriodEnd, previousFee);
  const newAmount = calculateSegment(effectiveDate, toDate, newFee);

  return formatAmountValue(oldAmount + newAmount);
}

function getRecurringBillingCycle(
  record,
  billingRecords = [],
  stage3Records = [],
  commonComplexAllocations = {},
  asOfDate = new Date(),
) {
  const latestBillingRecord = getLatestSubmittedBillingRecord(
    record,
    billingRecords,
  );

  const latestFirstTimePeriodTo = getLatestSubmittedFirstTimePeriodTo(record);
  const latestBillingPeriodTo = parseLocalDateValue(
    latestBillingRecord?.billingPeriodTo,
  );
  const latestPeriodTo =
    latestFirstTimePeriodTo && latestBillingPeriodTo
      ? new Date(
          Math.max(
            latestFirstTimePeriodTo.getTime(),
            latestBillingPeriodTo.getTime(),
          ),
        )
      : latestFirstTimePeriodTo || latestBillingPeriodTo;

  if (!latestPeriodTo) {
    return {
      periodFrom: "",
      periodTo: "",
      billingMonth: "",
      monthKey: "",
      actualSubscriptionFee: "",
      billingSubscriptionFee: "",
      eligibilityStatus: "First Billing Pending",
      eligible: false,
      currentMonthDue: false,
      alreadyProcessed: false,
      billingWorkStatus: "Pending",
      overdue: false,
    };
  }

  const today = getBillingAsOfDate(asOfDate) || new Date();
  const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const currentMonthStart = new Date(today.getFullYear(), today.getMonth(), 1);

  const normalizedSubscriptionType = normalizeValue(record?.subscriptionType)
    .toLowerCase();
  const normalizedSubscriptionMode = normalizeValue(record?.subscriptionMode)
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const isMonthly = ["", "monthly", "month"].includes(
    normalizedSubscriptionMode,
  );
  const isVariableMonthly =
    isMonthly && normalizedSubscriptionType === "variable";

  // Always move to the oldest cycle that has not yet been processed.
  // This keeps overdue/unbilled backlog ahead of the current cycle.
  const cycleStartDate = new Date(latestPeriodTo);
  cycleStartDate.setDate(cycleStartDate.getDate() + 1);

  const rawPeriodFrom = formatLocalDateValue(cycleStartDate);
  const rawPeriodTo = isMonthly
    ? formatLocalDateValue(
        new Date(cycleStartDate.getFullYear(), cycleStartDate.getMonth() + 1, 0),
      )
    : calculateBillingPeriodTo(rawPeriodFrom, record);

  const lifecyclePeriod = isMonthly
    ? getLifecycleAdjustedMonthlyPeriod(record, rawPeriodFrom, rawPeriodTo)
    : {
        periodFrom: rawPeriodFrom,
        periodTo: rawPeriodTo,
        billable: Boolean(rawPeriodFrom && rawPeriodTo),
      };

  const periodFrom = lifecyclePeriod.billable
    ? lifecyclePeriod.periodFrom
    : "";
  const periodTo = lifecyclePeriod.billable
    ? lifecyclePeriod.periodTo
    : "";

  const periodFromDate = parseLocalDateValue(periodFrom);
  const periodToDate = parseLocalDateValue(periodTo);
  const monthKey = getBillingMonthKey(cycleStartDate);
  const sourceRecordId = normalizeValue(record?.recordId);
  const screenCode = normalizeValue(record?.screenCode);

  const alreadyProcessed = billingRecords.some((billingRecord) => {
    const submitted = ["Submitted to Billing Team", "Submitted"].includes(
      normalizeValue(billingRecord?.submissionStatus),
    );
    if (!submitted || !isRecurringBillingRecord(billingRecord)) return false;

    const sameSource =
      sourceRecordId &&
      normalizeValue(billingRecord?.sourceRecordId) === sourceRecordId;
    const sameScreen =
      screenCode &&
      normalizeValue(billingRecord?.screenCode) === screenCode;

    if (!sameSource && !sameScreen) return false;

    return (
      normalizeValue(billingRecord?.billingPeriodFrom) === periodFrom &&
      normalizeValue(billingRecord?.billingPeriodTo) === periodTo
    );
  });

  const isCommonComplex = isCommonComplexBillingRecord(record);
  const savedAllocation = isCommonComplex
    ? getSavedSiteBillingAllocation(record, commonComplexAllocations)
    : "";

  const currentSiteWiseFee = isCommonComplex
    ? savedAllocation
    : normalizeValue(record?.priceChangeOverrideFee) ||
      resolveSiteWiseSubscriptionFee(record) ||
      normalizeValue(record?.subscriptionFee);

  const actualSubscriptionFee =
    normalizeValue(currentSiteWiseFee) ||
    normalizeValue(latestBillingRecord?.actualSubscriptionFee);

  const fixedRecurringBillingAmount =
    normalizeValue(currentSiteWiseFee) ||
    normalizeValue(latestBillingRecord?.actualSubscriptionFee) ||
    normalizeValue(record?.subscriptionFee);

  const billingSubscriptionFee = isVariableMonthly
    ? normalizeValue(record?.monthlyBillingAmounts?.[monthKey])
    : periodFrom && periodTo
      ? getPriceChangeAdjustedFixedAmount(
          record,
          periodFrom,
          periodTo,
          fixedRecurringBillingAmount,
        )
      : "";

  const resumeDate = getBillingResumeDate(record);
  const resumeAfterCycle =
    Boolean(resumeDate) && periodToDate && resumeDate > periodToDate;

  const cycleHasStarted =
    Boolean(periodFromDate) && periodFromDate <= todayStart;
  const billingMonthAvailable = isNormalBillingMonthAvailable(
    cycleStartDate,
    today,
  );
  const finalPeriodException = isFinalLifecycleBillingPeriod(
    record,
    periodFrom,
    periodTo,
    today,
  );
  const currentMonthDue =
    cycleHasStarted &&
    (billingMonthAvailable || finalPeriodException) &&
    !alreadyProcessed &&
    Boolean(periodFrom) &&
    Boolean(periodTo) &&
    !resumeAfterCycle;

  const overdue =
    currentMonthDue &&
    Boolean(periodToDate) &&
    periodToDate < currentMonthStart;

  let eligibilityStatus = resumeAfterCycle
    ? "Inactive / Paused"
    : overdue
      ? "Overdue / Unbilled"
      : currentMonthDue
        ? "Ready to Process"
        : "Future Billing";

  if (currentMonthDue && isVariableMonthly && !billingSubscriptionFee) {
    eligibilityStatus = "Amount Required";
  }

  const priceChangeBlock = getPriceChangeBillingBlock(
    record,
    stage3Records,
    periodFrom,
    periodTo,
  );

  if (priceChangeBlock.blocked && cycleHasStarted) {
    eligibilityStatus = "Price Change & Reallocation Pending";
  }

  return {
    periodFrom,
    periodTo,
    billingMonth: isMonthly
      ? getBillingMonthLabel(cycleStartDate)
      : formatBillingPeriod(periodFrom, periodTo),
    monthKey,
    actualSubscriptionFee,
    billingSubscriptionFee,
    eligibilityStatus,
    eligible: currentMonthDue && !priceChangeBlock.blocked,
    currentMonthDue,
    alreadyProcessed,
    billingWorkStatus: currentMonthDue ? "Unbilled" : alreadyProcessed ? "Billed" : "Future",
    overdue,
    priceChangeBlocked: priceChangeBlock.blocked,
    priceChangeBlockScope: priceChangeBlock.scope,
    priceChangeBlockLabel: priceChangeBlock.label,
  };
}

function getBillingTeamActualSubscriptionFee(
  record,
  stage3Records = [],
  commonComplexAllocations = {},
) {
  if (!record) {
    return "";
  }

  if (isVariableSubscriptionRecord(record)) {
    return "";
  }

  if (isCommonComplexBillingRecord(record)) {
    const allocationComplete =
      getCommonComplexBillingValidationError(
        record,
        stage3Records,
        commonComplexAllocations,
      ) === "";

    if (!allocationComplete) {
      return "";
    }

    return getSavedSiteBillingAllocation(record, commonComplexAllocations);
  }

  return (
    resolveSiteWiseSubscriptionFee(record) ||
    normalizeValue(record?.subscriptionFee)
  );
}

function getFirstTimeBillingTeamAmount(
  record,
  stage3Records = [],
  commonComplexAllocations = {},
) {
  if (!record) {
    return "";
  }

  if (isCommonComplexBillingRecord(record)) {
    const validationError = getCommonComplexBillingValidationError(
      record,
      stage3Records,
      commonComplexAllocations,
    );

    if (validationError) {
      return "";
    }

    const rows = getCommonComplexBillingRows(
      record,
      stage3Records,
      commonComplexAllocations,
    );
    const currentSiteKey = getCommonComplexBillingSiteKey(record.screenCode);
    const matchedRow = rows.find((row) => row.siteKey === currentSiteKey);

    return normalizeValue(matchedRow?.currentBillingCharge);
  }

  // Site / Screen-wise pricing stores the approved fee inside pricingGroups.
  // Use the resolved fee for first-time billing instead of depending only on
  // the top-level subscriptionFee field, which can be blank for grouped pricing.
  const resolvedSiteWiseFee =
    resolveSiteWiseSubscriptionFee(record) ||
    normalizeValue(record?.subscriptionFee);

  const recordWithEffectiveSiteFee = resolvedSiteWiseFee
    ? {
        ...record,
        subscriptionFee: resolvedSiteWiseFee,
      }
    : record;

  return (
    getStandaloneCurrentBillingCharge(recordWithEffectiveSiteFee) ||
    normalizeValue(record?.invoiceAmount) ||
    resolvedSiteWiseFee
  );
}

function isFirstTimeBillingTeamDownloadEligible(
  record,
  stage3Records = [],
  commonComplexAllocations = {},
) {
  if (!record || !isPendingFirstTimeBillingRecord(record)) {
    return false;
  }

  if (getBillingValidationErrors(record).length > 0) {
    return false;
  }

  if (
    getCommonComplexBillingValidationError(
      record,
      stage3Records,
      commonComplexAllocations,
    )
  ) {
    return false;
  }

  const normalizedSubscriptionType = normalizeValue(
    record?.subscriptionType,
  ).toLowerCase();
  const normalizedSubscriptionMode = normalizeValue(
    record?.subscriptionMode,
  )
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const isVariableMonthly =
    normalizedSubscriptionType === "variable" &&
    ["", "monthly", "month"].includes(normalizedSubscriptionMode);

  if (isVariableMonthly) {
    const firstBillingRow = buildMonthlyBillingRows(record).find(
      (row) => row.invoiceEligible === true,
    );

    return Boolean(
      firstBillingRow && normalizeValue(firstBillingRow.amount),
    );
  }

  const actualSubscriptionFee = getBillingTeamActualSubscriptionFee(
    record,
    stage3Records,
    commonComplexAllocations,
  );
  const billingSubscriptionFee = getFirstTimeBillingTeamAmount(
    record,
    stage3Records,
    commonComplexAllocations,
  );

  return Boolean(actualSubscriptionFee && billingSubscriptionFee);
}

function buildBillingTeamExportRow(
  record,
  billingSubscriptionFee,
  actualSubscriptionFee,
  {
    periodFrom = "",
    periodTo = "",
    remarks = "Subscription",
  } = {},
) {
  return {
    "Screen Code": normalizeValue(record?.screenCode),
    "Screen Name": normalizeValue(record?.screenName),
    Location: normalizeValue(record?.location),
    State: normalizeValue(record?.state),
    "Billing Code": normalizeValue(record?.billingCode),
    "Billing Name": normalizeValue(record?.billingName),
    "Billing State": normalizeValue(record?.state),
    "Pre Tax Amount": normalizeValue(actualSubscriptionFee),
    "Billing Subscription Fee (Pre Tax Amount)": normalizeValue(
      billingSubscriptionFee,
    ),
    "Billing Period From": normalizeValue(periodFrom),
    "Billing Period To": normalizeValue(periodTo),
    "Subscription Mode": normalizeValue(record?.subscriptionMode),
    Remarks: normalizeValue(remarks) || "Subscription",
  };
}

function buildRecurringBillingTeamExportRow(record) {
  const recurringCycle = record?.recurringCycle || {};

  return buildBillingTeamExportRow(
    record,
    recurringCycle?.billingSubscriptionFee,
    recurringCycle?.actualSubscriptionFee,
    {
      periodFrom: recurringCycle?.periodFrom,
      periodTo: recurringCycle?.periodTo,
      remarks: "Subscription",
    },
  );
}

function buildFirstTimeBillingTeamExportRows(
  record,
  stage3Records = [],
  commonComplexAllocations = {},
) {
  if (!record) {
    return [];
  }

  const actualSubscriptionFee = getBillingTeamActualSubscriptionFee(
    record,
    stage3Records,
    commonComplexAllocations,
  );
  const monthlyRows = buildMonthlyBillingRows(
    record,
    actualSubscriptionFee,
  );

  if (monthlyRows.length > 0) {
    const firstBillingRow = monthlyRows[0];

    return [
      buildBillingTeamExportRow(
        record,
        firstBillingRow.amount,
        actualSubscriptionFee,
        {
          periodFrom: firstBillingRow.periodFrom,
          periodTo: firstBillingRow.periodTo,
          remarks: "Subscription",
        },
      ),
    ];
  }

  const billingSubscriptionFee = getFirstTimeBillingTeamAmount(
    record,
    stage3Records,
    commonComplexAllocations,
  );

  return [
    buildBillingTeamExportRow(
      record,
      billingSubscriptionFee,
      actualSubscriptionFee,
      {
        periodFrom: record?.billingPeriodFrom,
        periodTo: record?.billingPeriodTo,
        remarks: "Subscription",
      },
    ),
  ];
}

function buildFirstTimeBillingTeamDownloadRows(
  records = [],
  stage3Records = [],
  commonComplexAllocations = {},
) {
  const exportRows = [];
  const processedCommonGroups = new Set();

  records.forEach((record) => {
    if (!record) return;

    const canCombine =
      isCommonComplexBillingRecord(record) &&
      canUseCombinedCommonBilling(
        record,
        stage3Records,
        commonComplexAllocations,
      );

    if (!canCombine) {
      exportRows.push(
        ...buildFirstTimeBillingTeamExportRows(
          record,
          stage3Records,
          commonComplexAllocations,
        ),
      );
      return;
    }

    const groupKey = getCommonComplexBillingGroupKey(record);
    if (!groupKey || processedCommonGroups.has(groupKey)) {
      return;
    }
    processedCommonGroups.add(groupKey);

    const commonRows = getCommonComplexBillingRows(
      record,
      stage3Records,
      commonComplexAllocations,
    );
    const relatedRecords = commonRows
      .map((row) => row.relatedRecord)
      .filter(Boolean);

    const childExportRows = relatedRecords.flatMap((relatedRecord) =>
      buildFirstTimeBillingTeamExportRows(
        relatedRecord,
        stage3Records,
        commonComplexAllocations,
      ),
    );

    if (childExportRows.length === 0) {
      return;
    }

    const representative = relatedRecords[0] || record;
    const commonSubscriptionFee = resolveCommonComplexSubscriptionFee(
      representative,
      stage3Records,
      commonComplexAllocations,
    );
    const combinedBillingAmount = childExportRows.reduce((sum, row) => {
      const amount = parseAmountValue(
        row["Billing Subscription Fee (Pre Tax Amount)"],
      );
      return sum + (amount === null ? 0 : amount);
    }, 0);

    const periodFromDates = childExportRows
      .map((row) => parseLocalDateValue(row["Billing Period From"]))
      .filter(Boolean);
    const periodToDates = childExportRows
      .map((row) => parseLocalDateValue(row["Billing Period To"]))
      .filter(Boolean);
    const locations = Array.from(
      new Set(
        relatedRecords
          .map((row) => normalizeValue(row.location))
          .filter(Boolean),
      ),
    );
    const states = Array.from(
      new Set(
        relatedRecords
          .map((row) => normalizeValue(row.state))
          .filter(Boolean),
      ),
    );

    const combinedScreenCodes = Array.from(
      new Set(
        relatedRecords
          .map((row) => normalizeValue(row?.screenCode))
          .filter(Boolean),
      ),
    );
    const combinedScreenNames = Array.from(
      new Set(
        relatedRecords
          .map((row) => normalizeValue(row?.screenName))
          .filter(Boolean),
      ),
    );

    exportRows.push({
      "Screen Code": combinedScreenCodes.join(", "),
      "Screen Name": combinedScreenNames.join(", "),
      Location: locations.length <= 1 ? locations[0] || "" : "Multiple",
      State: states.length <= 1 ? states[0] || "" : "Multiple",
      "Billing Code": normalizeValue(representative?.billingCode),
      "Billing Name": normalizeValue(representative?.billingName),
      "Billing State": normalizeValue(representative?.state),
      "Pre Tax Amount": normalizeValue(commonSubscriptionFee),
      "Billing Subscription Fee (Pre Tax Amount)": formatAmountValue(
        combinedBillingAmount,
      ),
      "Billing Period From": periodFromDates.length
        ? formatLocalDateValue(
            new Date(
              Math.min(...periodFromDates.map((date) => date.getTime())),
            ),
          )
        : "",
      "Billing Period To": periodToDates.length
        ? formatLocalDateValue(
            new Date(
              Math.max(...periodToDates.map((date) => date.getTime())),
            ),
          )
        : "",
      "Subscription Mode": normalizeValue(representative?.subscriptionMode),
      Remarks: "Subscription",
    });
  });

  return exportRows;
}


function getBillingTransactionKey(record) {
  return [
    normalizeValue(record?.screenCode || record?.siteScope).toUpperCase(),
    normalizeValue(record?.billingPeriodFrom),
    normalizeValue(record?.billingPeriodTo),
  ]
    .filter(Boolean)
    .join("::");
}

function normalizeInvoiceReference(value) {
  return normalizeValue(value).toUpperCase();
}

function getInvoiceTransactionGroup(record, entry = {}) {
  const explicitGroup = normalizeValue(
      entry.invoiceTransactionGroup ||
      entry.commonInvoiceGroup ||
      record?.invoiceTransactionGroup ||
      record?.commonInvoiceGroup ||
      record?.otfTransactionGroupKey ||
      record?.transactionId ||
      record?.stage1GroupId,
  );

  if (explicitGroup) return explicitGroup.toUpperCase();

  if (isCommonComplexBillingRecord(record)) {
    return [
      normalizeValue(record?.billingCode),
      normalizeValue(record?.complexCode),
      "COMMON",
    ]
      .join("::")
      .toUpperCase();
  }

  return "";
}

function getInvoicePeriodKey(entry = {}) {
  return [
    normalizeValue(entry.periodFrom),
    normalizeValue(entry.periodTo),
  ].join("::");
}

function findInvoiceConflict({
  candidateRecord,
  candidateEntry,
  billingRecords = [],
  ignoredBillingRecordId = "",
}) {
  const candidateInvoice = normalizeInvoiceReference(candidateEntry?.invoiceNumber);
  const candidateBillingCode = normalizeInvoiceReference(candidateRecord?.billingCode);

  if (!candidateInvoice || !candidateBillingCode) return "";

  const candidateGroup = getInvoiceTransactionGroup(candidateRecord, candidateEntry);
  const candidatePeriod = getInvoicePeriodKey(candidateEntry);

  for (const record of billingRecords) {
    if (
      normalizeInvoiceReference(record?.billingCode) !== candidateBillingCode
    ) {
      continue;
    }

    for (const entry of buildBillingInvoicePeriods(record)) {
      if (
        ignoredBillingRecordId &&
        normalizeValue(record?.billingRecordId) === ignoredBillingRecordId &&
        normalizeValue(entry?.entryId) === normalizeValue(candidateEntry?.entryId)
      ) {
        continue;
      }

      if (normalizeInvoiceReference(entry?.invoiceNumber) !== candidateInvoice) {
        continue;
      }

      const sameLegitimateGroup =
        candidateGroup &&
        candidateGroup === getInvoiceTransactionGroup(record, entry) &&
        candidatePeriod === getInvoicePeriodKey(entry);

      if (!sameLegitimateGroup) {
        return `Invoice Number ${candidateEntry.invoiceNumber} has already been used for this Billing Code.`;
      }
    }
  }

  return "";
}

function isRecurringBillingRecord(record) {
  return Boolean(
    normalizeValue(record?.billingMonth) ||
      normalizeValue(record?.billingCycleType).toLowerCase() === "recurring",
  );
}

function getFirstTimeBillingSiteKey(record) {
  return (
    normalizeValue(record?.backendSiteId) ||
    normalizeValue(record?.screenCode || record?.siteScope).toUpperCase()
  );
}

function getBillingRecordIdentity(record) {
  return getBillingTransactionKey(record);
}

function dedupeBillingRecords(records = []) {
  const seenFirstTimeSites = new Set();
  const seenRecurringTransactions = new Set();

  const sortedRecords = records.slice().sort((left, right) => {
    const leftTime =
      new Date(left?.submittedAt || left?.createdAt || left?.lastUpdatedAt || 0)
        .getTime() || 0;
    const rightTime =
      new Date(right?.submittedAt || right?.createdAt || right?.lastUpdatedAt || 0)
        .getTime() || 0;

    // Keep the earliest permanent first-time Billing Record.
    return leftTime - rightTime;
  });

  return sortedRecords.filter((record) => {
    if (isRecurringBillingRecord(record)) {
      const recurringKey =
        getBillingTransactionKey(record) || record?.billingRecordId;

      if (!recurringKey) {
        return true;
      }

      if (seenRecurringTransactions.has(recurringKey)) {
        return false;
      }

      seenRecurringTransactions.add(recurringKey);
      return true;
    }

    const firstTimeSiteKey = getFirstTimeBillingSiteKey(record);

    if (!firstTimeSiteKey) {
      return true;
    }

    if (seenFirstTimeSites.has(firstTimeSiteKey)) {
      return false;
    }

    seenFirstTimeSites.add(firstTimeSiteKey);
    return true;
  });
}

function buildBillingInvoicePeriods(record) {
  if (!record) {
    return [];
  }

  const periodFrom = parseLocalDateValue(record.billingPeriodFrom);
  const periodTo = parseLocalDateValue(record.billingPeriodTo);

  if (!periodFrom || !periodTo || periodTo < periodFrom) {
    return [];
  }

  const submittedFirstTimeRows = buildSubmittedFirstTimePeriodRows(record);
  if (submittedFirstTimeRows.length) {
    return submittedFirstTimeRows;
  }

  const frequency = getBillingPeriodFrequency(record);
  const savedEntries = Array.isArray(record.invoiceEntries)
    ? record.invoiceEntries
    : [];

  if (frequency !== "month") {
    const existing = savedEntries[0] || {};

    return [
      {
        entryId: existing.entryId || getBillingTransactionKey(record) || "billing-period",
        billingMonth: formatBillingPeriod(
          record.billingPeriodFrom,
          record.billingPeriodTo,
        ),
        periodFrom: normalizeValue(record.billingPeriodFrom),
        periodTo: normalizeValue(record.billingPeriodTo),
        invoiceNumber:
          normalizeValue(existing.invoiceNumber) ||
          normalizeValue(record.invoiceNumber),
        invoiceDate:
          normalizeValue(existing.invoiceDate) ||
          normalizeValue(record.invoiceDate),
        invoiceAmountBeforeGST:
          normalizeValue(existing.invoiceAmountBeforeGST) ||
          normalizeValue(record.invoiceAmountBeforeGST),
        invoiceStatus:
          normalizeValue(existing.invoiceStatus) ||
          getBillingRecordInvoiceStatus(existing.invoiceNumber ? existing : record),
        paymentStatus:
          normalizeValue(existing.paymentStatus) ||
          normalizeValue(record.paymentStatus) ||
          "Pending",
        paymentReceivedDate:
          normalizeValue(existing.paymentReceivedDate) ||
          normalizeValue(record.paymentReceivedDate),
        paymentRemarks:
          normalizeValue(existing.paymentRemarks) ||
          normalizeValue(record.paymentRemarks),
        receiptNumber:
          normalizeValue(existing.receiptNumber) ||
          normalizeValue(record.receiptNumber) ||
          normalizeValue(record.paymentReceiptNumber) ||
          normalizeValue(record.erpReceiptNumber),
        receiptDate:
          normalizeValue(existing.receiptDate) ||
          normalizeValue(record.receiptDate) ||
          normalizeValue(record.paymentReceiptDate) ||
          normalizeValue(record.erpReceiptDate),
        receivedAmount:
          normalizeValue(existing.receivedAmount) ||
          normalizeValue(record.receivedAmount) ||
          normalizeValue(record.paymentReceivedAmount) ||
          normalizeValue(record.erpReceivedAmount),
      },
    ];
  }

  const calculatedMonthlyRows = buildMonthlyBillingRows(record);
  const calculatedByMonth = new Map(
    calculatedMonthlyRows.map((row) => [row.monthKey, row]),
  );
  const savedByMonth = new Map(
    savedEntries.map((entry) => [normalizeValue(entry.entryId), entry]),
  );

  const rows = [];
  let cursor = new Date(periodFrom.getFullYear(), periodFrom.getMonth(), 1);
  const lastMonth = new Date(periodTo.getFullYear(), periodTo.getMonth(), 1);

  while (cursor <= lastMonth) {
    const year = cursor.getFullYear();
    const month = cursor.getMonth();
    const monthKey = getBillingMonthKey(cursor);
    const monthStart = new Date(year, month, 1);
    const monthEnd = new Date(year, month + 1, 0);
    const effectiveFrom =
      year === periodFrom.getFullYear() && month === periodFrom.getMonth()
        ? periodFrom
        : monthStart;
    const effectiveTo =
      year === periodTo.getFullYear() && month === periodTo.getMonth()
        ? periodTo
        : monthEnd;
    const existing = savedByMonth.get(monthKey) || {};
    const calculated = calculatedByMonth.get(monthKey) || {};
    const isFirstMonth = rows.length === 0;

    rows.push({
      entryId: monthKey,
      billingMonth: getBillingMonthLabel(cursor),
      periodFrom: formatLocalDateValue(effectiveFrom),
      periodTo: formatLocalDateValue(effectiveTo),
      invoiceNumber:
        normalizeValue(existing.invoiceNumber) ||
        (savedEntries.length === 0 && isFirstMonth
          ? normalizeValue(record.invoiceNumber)
          : ""),
      invoiceDate:
        normalizeValue(existing.invoiceDate) ||
        (savedEntries.length === 0 && isFirstMonth
          ? normalizeValue(record.invoiceDate)
          : ""),
      invoiceAmountBeforeGST:
        normalizeValue(existing.invoiceAmountBeforeGST) ||
        normalizeValue(calculated.amount) ||
        (isFirstMonth ? normalizeValue(record.billingAmountBeforeGST) : ""),
      invoiceStatus:
        normalizeValue(existing.invoiceStatus) ||
        getBillingRecordInvoiceStatus(existing),
      paymentStatus:
        normalizeValue(existing.paymentStatus) ||
        (savedEntries.length === 0 && isFirstMonth
          ? normalizeValue(record.paymentStatus) || "Pending"
          : "Pending"),
      paymentReceivedDate:
        normalizeValue(existing.paymentReceivedDate) ||
        (savedEntries.length === 0 && isFirstMonth
          ? normalizeValue(record.paymentReceivedDate)
          : ""),
      paymentRemarks:
        normalizeValue(existing.paymentRemarks) ||
        (savedEntries.length === 0 && isFirstMonth
          ? normalizeValue(record.paymentRemarks)
          : ""),
      receiptNumber:
        normalizeValue(existing.receiptNumber) ||
        (savedEntries.length === 0 && isFirstMonth
          ? normalizeValue(
              record.receiptNumber ||
                record.paymentReceiptNumber ||
                record.erpReceiptNumber,
            )
          : ""),
      receiptDate:
        normalizeValue(existing.receiptDate) ||
        (savedEntries.length === 0 && isFirstMonth
          ? normalizeValue(
              record.receiptDate ||
                record.paymentReceiptDate ||
                record.erpReceiptDate,
            )
          : ""),
      receivedAmount:
        normalizeValue(existing.receivedAmount) ||
        (savedEntries.length === 0 && isFirstMonth
          ? normalizeValue(
              record.receivedAmount ||
                record.paymentReceivedAmount ||
                record.erpReceivedAmount,
            )
          : ""),
    });

    cursor = new Date(year, month + 1, 1);
  }

  return rows;
}

function getBillingRecordInvoiceStatus(record) {
  const invoiceNumber = normalizeValue(record?.invoiceNumber);
  const invoiceDate = normalizeValue(record?.invoiceDate);
  const invoiceAmount = normalizeValue(record?.invoiceAmountBeforeGST);

  if (invoiceNumber || invoiceDate || invoiceAmount) {
    return "Invoiced";
  }

  return "Pending Invoice";
}

function buildBillingRecordSnapshot(
  record,
  stage3Records = [],
  commonComplexAllocations = {},
  submittedAt = "",
  submissionStatus = "Submitted to Billing Team",
) {
  if (!record) {
    return null;
  }

  const isCommonComplex = isCommonComplexBillingRecord(record);
  const commonComplexRows = isCommonComplex
    ? getCommonComplexBillingRows(
        record,
        stage3Records,
        commonComplexAllocations,
      )
    : [];
  const matchedCommonComplexRow = commonComplexRows.find(
    (row) => row.siteKey === getCommonComplexBillingSiteKey(record.screenCode),
  );
  const billingMode = isCommonComplex
    ? resolveCommonComplexBillingStatus(
        record,
        stage3Records,
        commonComplexAllocations,
      )
    : "Standalone / Site-wise";
  const configuredCommonSubscriptionFee = isCommonComplex
    ? resolveCommonComplexSubscriptionFee(
        record,
        stage3Records,
        commonComplexAllocations,
      )
    : "";
  const effectiveScreenFee = isCommonComplex
    ? normalizeValue(matchedCommonComplexRow?.allocatedFee) ||
      getSavedSiteBillingAllocation(record, commonComplexAllocations)
    : resolveSiteWiseSubscriptionFee(record) ||
      normalizeValue(record.subscriptionFee);
  const billingAmountBeforeGST =
    normalizeValue(record?.firstTimeBillingSummary?.actualBillingAmount) ||
    normalizeValue(matchedCommonComplexRow?.currentBillingCharge) ||
    getStandaloneCurrentBillingCharge(record) ||
    normalizeValue(record.invoiceAmount) ||
    normalizeValue(record.subscriptionFee);
  const siteScope =
    normalizeValue(matchedCommonComplexRow?.screenCode) ||
    normalizeValue(record.screenCode) ||
    normalizeValue(record.complexCode) ||
    "-";
  const billingRecordId = getBillingRecordIdentity(record);
  const invoiceAmountBeforeGST = normalizeValue(record.invoiceAmount);

  return {
    ...record,
    subscriptionFee:
      normalizeValue(effectiveScreenFee) || normalizeValue(record.subscriptionFee),
    effectiveScreenFee: normalizeValue(effectiveScreenFee),
    configuredCommonSubscriptionFee: normalizeValue(
      configuredCommonSubscriptionFee,
    ),
    billingRecordId,
    sourceRecordId: normalizeValue(record.recordId),
    billingCode: normalizeValue(record.billingCode),
    complexCode: normalizeValue(record.complexCode),
    siteScope,
    billingPeriodFrom: normalizeValue(record.billingPeriodFrom),
    billingPeriodTo: normalizeValue(record.billingPeriodTo),
    billingAmountBeforeGST,
    billingMode,
    submissionStatus,
    submittedAt,
    createdAt: submittedAt,
    lastUpdatedAt: submittedAt,
    billingCommercialLocked: submissionStatus === "Submitted to Billing Team",
    billingCommercialLockedAt:
      submissionStatus === "Submitted to Billing Team" ? submittedAt : "",
    billingCommercialLockReason:
      submissionStatus === "Submitted to Billing Team"
        ? "Data shared with Billing Team"
        : "",
    auditTrail:
      record?.firstTimeWaiver?.applicable === "Yes"
        ? [
            createBillingAuditEntry({
              event: "First Time Billing Waiver Applied",
              field: "firstTimeWaiver",
              previousValue: "No waiver",
              newValue: JSON.stringify(record.firstTimeWaiver),
              remarks: normalizeValue(record.firstTimeWaiver.reason),
              timestamp: submittedAt,
            }),
          ]
        : [],
    billingCommercialLockSnapshot: {
      installationDate: normalizeValue(record.installationDate),
      liveDate: normalizeValue(record.liveDate),
      trialPeriod: normalizeValue(record.trialPeriod),
      trialPeriodExtension: normalizeValue(
        record.trialPeriodExtension ||
          record.totalTrialPeriodExtension ||
          record.trialExtension,
      ),
      billingStartDate: normalizeValue(record.billingStartDate),
      mouStartDate: normalizeValue(record.mouStartDate),
      mouEndDate: normalizeValue(record.mouEndDate),
      subscriptionType: normalizeValue(record.subscriptionType),
      subscriptionMode: normalizeValue(record.subscriptionMode),
      subscriptionFee:
        normalizeValue(effectiveScreenFee) ||
        normalizeValue(record.subscriptionFee),
      configuredCommonSubscriptionFee: normalizeValue(
        configuredCommonSubscriptionFee,
      ),
      otfApplicable: normalizeValue(record.otfApplicable),
      otfType: normalizeValue(record.otfType),
      otfAmount: normalizeValue(record.otfAmount),
      pricingMethod: normalizeValue(record.pricingMethod),
      pricingGroups: Array.isArray(record.pricingGroups)
        ? record.pricingGroups.map((group) => ({ ...group }))
        : [],
    },
    invoiceNumber: normalizeValue(record.invoiceNumber),
    invoiceDate: normalizeValue(record.invoiceDate),
    invoiceAmountBeforeGST,
    invoiceStatus: getBillingRecordInvoiceStatus(record),
    paymentStatus: normalizeValue(record.paymentStatus) || "Pending",
    paymentReceivedDate: normalizeValue(record.paymentReceivedDate),
    paymentRemarks: normalizeValue(record.paymentRemarks),
    commonComplexBillingRows: commonComplexRows,
    commonComplexBillingValidationError: isCommonComplex
      ? getCommonComplexBillingValidationError(
          record,
          stage3Records,
          commonComplexAllocations,
        )
      : "",
  };
}

function getBillingRecordEffectiveScreenFee(
  record,
  commonComplexAllocations = {},
) {
  if (!record) {
    return "";
  }

  if (!isCommonComplexBillingRecord(record)) {
    return (
      normalizeValue(record.effectiveScreenFee) ||
      resolveSiteWiseSubscriptionFee(record) ||
      normalizeValue(record.subscriptionFee)
    );
  }

  const screenCode = getCommonComplexBillingSiteKey(
    record.screenCode || record.siteScope,
  );
  const snapshotRows = Array.isArray(record.commonComplexBillingRows)
    ? record.commonComplexBillingRows
    : [];
  const snapshotRow = snapshotRows.find(
    (row) => getCommonComplexBillingSiteKey(row?.screenCode) === screenCode,
  );

  return (
    normalizeValue(record.effectiveScreenFee) ||
    normalizeValue(snapshotRow?.allocatedFee) ||
    getSavedSiteBillingAllocation(record, commonComplexAllocations) ||
    normalizeValue(record.actualSubscriptionFee) ||
    normalizeValue(record.subscriptionFee)
  );
}

function getFirstTimeDisplaySubscriptionFee(
  record,
  billingRecords = [],
  stage3Records = [],
  commonComplexAllocations = {},
) {
  const historicalRecord =
    getFirstSubmittedBillingRecord(record, billingRecords) ||
    (record?.firstTimeBillingRecord &&
    typeof record.firstTimeBillingRecord === "object"
      ? record.firstTimeBillingRecord
      : null);

  if (
    getPersistedFirstTimeBillingStatus(record) === "Completed" &&
    historicalRecord
  ) {
    return getBillingRecordEffectiveScreenFee(
      historicalRecord,
      commonComplexAllocations,
    );
  }

  return getBillingTeamActualSubscriptionFee(
    record,
    stage3Records,
    commonComplexAllocations,
  );
}

function formatBillingPeriod(from, to) {
  const start = normalizeValue(from);
  const end = normalizeValue(to);

  if (start && end) {
    return `${start} to ${end}`;
  }

  return start || end || "-";
}

function getDerivedBillingTreatment(record) {
  const existingBillingTreatment = normalizeValue(record.billingTreatment);
  if (existingBillingTreatment) {
    return existingBillingTreatment;
  }

  const billingStartDate = normalizeValue(record.billingStartDate);
  if (!isValidDateValue(billingStartDate)) {
    return "";
  }

  const day = Number(billingStartDate.slice(8, 10));
  return day === 1 ? "Full Month" : "Pro-rata";
}

function getPersistedFirstTimeBillingStatus(record) {
  const status = normalizeValue(record?.firstTimeBillingValidationStatus)
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (status === "completed") return "Completed";
  if (status === "validated") return "Validated";
  return "Pending";
}

const firstTimeBillingImpactingFields = new Set([
  "subscriptionFee",
  "subscriptionType",
  "subscriptionMode",
  "billingStartDate",
  "billingPeriodFrom",
  "billingPeriodTo",
  "trialPeriod",
  "trialPeriodExtension",
  "monthlyBillingAmounts",
  "pricingMethod",
  "pricingGroups",
  "allocationRows",
  "commercialApplicable",
  "foc",
  "otfApplicable",
  "firstTimeWaiver",
  "firstTimePeriodBreakdown",
]);

function isCompletedFirstTimeBillingRecord(record) {
  return (
    getPersistedFirstTimeBillingStatus(record) === "Completed" ||
    normalizeValue(record?.billingVerificationStatus).toLowerCase() ===
      "submitted to billing team" ||
    Boolean(record?.firstTimeBillingRecord?.billingRecordId)
  );
}

function invalidateFirstTimeBillingValidation(record, changedField) {
  if (
    !record ||
    !firstTimeBillingImpactingFields.has(changedField) ||
    isCompletedFirstTimeBillingRecord(record)
  ) {
    return record;
  }

  return {
    ...record,
    firstTimeBillingValidationStatus: "Billing Calculation",
    firstTimeBillingValidatedAt: "",
    firstTimeBillingValidatedBy: "",
  };
}

function isPendingFirstTimeBillingRecord(record) {
  if (getPersistedFirstTimeBillingStatus(record) === "Completed") {
    return false;
  }

  const verificationStatus = normalizeValue(record.billingVerificationStatus);
  const isHumanErrorCorrection =
    normalizeValue(record?.correctionType) === "Human Error";
  const correctionRequestedAt = normalizeValue(record?.correctionRequestedAt);
  const correctionReverifiedAt = normalizeValue(record?.correctionReverifiedAt);

  if (
    isHumanErrorCorrection &&
    correctionRequestedAt &&
    !correctionReverifiedAt
  ) {
    return true;
  }

  return (
    !verificationStatus ||
    verificationStatus === "First Billing Pending Approval"
  );
}


function isFirstTimeBillingSiteValidatedForCurrentCycle(record) {
  const validatedAt = normalizeValue(record?.firstTimeBillingValidatedAt);

  if (!validatedAt) {
    return false;
  }

  const correctionRequestedAt = normalizeValue(record?.correctionRequestedAt);

  if (!correctionRequestedAt) {
    return true;
  }

  const validatedTime = new Date(validatedAt).getTime() || 0;
  const correctionTime = new Date(correctionRequestedAt).getTime() || 0;

  return validatedTime >= correctionTime;
}

function isFirstTimeBillingSiteReadyForSubmission(
  record,
  stage3Records = [],
  commonComplexAllocations = {},
) {
  if (!record || !isPendingFirstTimeBillingRecord(record)) {
    return false;
  }

  if (!isFirstTimeBillingSiteValidatedForCurrentCycle(record)) {
    return false;
  }

  if (getBillingValidationErrors(record).length > 0) {
    return false;
  }

  if (
    getCommonComplexBillingValidationError(
      record,
      stage3Records,
      commonComplexAllocations,
    )
  ) {
    return false;
  }

  const normalizedSubscriptionType = normalizeValue(
    record?.subscriptionType,
  ).toLowerCase();
  const normalizedSubscriptionMode = normalizeValue(
    record?.subscriptionMode,
  )
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const isVariableMonthly =
    normalizedSubscriptionType === "variable" &&
    ["", "monthly", "month"].includes(normalizedSubscriptionMode);

  // Variable monthly first billing covers only the first billing cycle.
  // Later monthly cycles belong to Recurring Billing.
  if (isVariableMonthly) {
    const firstBillingRow = buildMonthlyBillingRows(record).find(
      (row) => row.invoiceEligible === true,
    );

    return Boolean(
      firstBillingRow && normalizeValue(firstBillingRow.amount),
    );
  }

  return isFirstTimeBillingTeamDownloadEligible(
    record,
    stage3Records,
    commonComplexAllocations,
  );
}

function getFirstSubmittedBillingRecord(record, billingRecords = []) {
  const sourceRecordId = normalizeValue(record?.recordId);
  const screenCode = normalizeValue(record?.screenCode);

  const matchingSubmittedRecords = billingRecords
    .filter((billingRecord) => {
      const submitted = ["Submitted to Billing Team", "Submitted"].includes(
        normalizeValue(billingRecord?.submissionStatus),
      );

      if (!submitted) {
        return false;
      }

      const sameSource =
        sourceRecordId &&
        normalizeValue(billingRecord?.sourceRecordId) === sourceRecordId;
      const sameScreen =
        screenCode &&
        normalizeValue(billingRecord?.screenCode) === screenCode;

      return sameSource || sameScreen;
    })
    .sort((left, right) => {
      // First-billing lifecycle movement is based on when the billing was
      // actually submitted/processed, not on the service billing-period dates.
      const leftDate =
        new Date(left?.submittedAt || left?.createdAt || 0).getTime() || 0;
      const rightDate =
        new Date(right?.submittedAt || right?.createdAt || 0).getTime() || 0;

      return leftDate - rightDate;
    });

  return matchingSubmittedRecords[0] || null;
}

function getLatestSubmittedFirstTimePeriodTo(record) {
  const periods = Array.isArray(record?.firstTimePeriodBreakdown)
    ? record.firstTimePeriodBreakdown
    : [];
  const dates = periods
    .map((period) => parseLocalDateValue(period?.periodTo))
    .filter(Boolean);

  if (!dates.length) {
    return null;
  }

  return new Date(Math.max(...dates.map((date) => date.getTime())));
}

function getFirstBillingCompletionDate(record, billingRecords = []) {
  const persistedCompletedAt = normalizeValue(
    record?.firstTimeBillingCompletedAt,
  );
  const persistedDate = persistedCompletedAt
    ? new Date(persistedCompletedAt)
    : null;

  if (persistedDate && !Number.isNaN(persistedDate.getTime())) {
    return persistedDate;
  }

  const firstBillingRecord = getFirstSubmittedBillingRecord(
    record,
    billingRecords,
  );
  const submittedAt =
    firstBillingRecord?.submittedAt || firstBillingRecord?.createdAt;
  const submittedDate = submittedAt ? new Date(submittedAt) : null;

  return submittedDate && !Number.isNaN(submittedDate.getTime())
    ? submittedDate
    : null;
}

function isFirstBillingSubmittedInCurrentMonth(record, billingRecords = []) {
  // A site stays in 3A through the month in which the first billing was
  // completed, even if the billing period itself starts in an earlier month.
  const firstBillingDate = getFirstBillingCompletionDate(record, billingRecords);

  if (!firstBillingDate || Number.isNaN(firstBillingDate.getTime())) {
    return false;
  }

  const today = new Date();

  return (
    firstBillingDate.getFullYear() === today.getFullYear() &&
    firstBillingDate.getMonth() === today.getMonth()
  );
}

function isFirstBillingBeforeCurrentMonth(record, billingRecords = []) {
  const firstBillingDate = getFirstBillingCompletionDate(record, billingRecords);

  if (!firstBillingDate || Number.isNaN(firstBillingDate.getTime())) {
    return false;
  }

  const today = new Date();
  const currentMonthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const firstBillingMonthStart = new Date(
    firstBillingDate.getFullYear(),
    firstBillingDate.getMonth(),
    1,
  );

  return firstBillingMonthStart < currentMonthStart;
}

function getStage1GroupKey(record) {
  return (
    normalizeValue(record?.stage1GroupId) ||
    normalizeValue(record?.savedAt) ||
    normalizeValue(record?.recordId) ||
    normalizeValue(record?.screenCode)
  );
}

function getOtfTransactionId(record, pricingMethod, groupKey) {
  if (normalizeValue(pricingMethod).toLowerCase() === "common") {
    return `OTF-COMMON-${groupKey}`;
  }

  return `OTF-SITE-${normalizeValue(record?.recordId) || normalizeValue(record?.screenCode) || groupKey}`;
}

function getCommercialLookupToken(value) {
  return normalizeValue(value).replace(/\s+/g, " ").toUpperCase();
}

function getStage1CommercialSiteLookups(record) {
  const stage1GroupRows = Array.isArray(record?.stage1GroupRows)
    ? record.stage1GroupRows
    : [];
  const siteIdLookup = new Map();
  const screenNameLookup = new Map();
  const aliasLookup = new Map();

  stage1GroupRows.forEach((row, index) => {
    const screenCode = getCommercialLookupToken(row?.screenCode);
    const screenName = getCommercialLookupToken(row?.screenName);
    const erpScreenName = getCommercialLookupToken(row?.erpScreenCode || row?.erpScreenName);
    const alias = `S${index + 1}`;
    const siteEntry = {
      screenCode,
      screenName,
      erpScreenName,
      alias,
    };

    if (screenCode) {
      siteIdLookup.set(screenCode, siteEntry);
    }

    if (screenName) {
      screenNameLookup.set(screenName, siteEntry);
    }

    if (erpScreenName) {
      screenNameLookup.set(erpScreenName, siteEntry);
    }

    aliasLookup.set(alias, siteEntry);
  });

  return {
    stage1GroupRows,
    siteIdLookup,
    screenNameLookup,
    aliasLookup,
  };
}

function resolvePricingGroupSiteIds(group, record, commercialLookups) {
  const resolvedSiteIds = Array.isArray(group?.resolvedApplicableSiteIds)
    ? group.resolvedApplicableSiteIds
        .map((screenCode) => normalizeValue(screenCode).toUpperCase())
        .filter(Boolean)
    : [];

  if (resolvedSiteIds.length > 0) {
    return Array.from(new Set(resolvedSiteIds));
  }

  const rawApplicableSites = normalizeValue(group?.applicableSites).replace(/\s+/g, " ");

  if (!rawApplicableSites) {
    return [];
  }

  const rawTokens = rawApplicableSites
    .split(/[,;]+/)
    .map((token) => getCommercialLookupToken(token))
    .filter(Boolean);

  const matchedSiteIds = [];
  const seenSiteIds = new Set();

  function addSiteId(screenCode) {
    const normalizedSiteId = getCommercialLookupToken(screenCode);
    if (!normalizedSiteId || seenSiteIds.has(normalizedSiteId)) {
      return;
    }

    seenSiteIds.add(normalizedSiteId);
    matchedSiteIds.push(normalizedSiteId);
  }

  rawTokens.forEach((token) => {
    if (commercialLookups.siteIdLookup.has(token)) {
      addSiteId(commercialLookups.siteIdLookup.get(token).screenCode);
      return;
    }

    if (commercialLookups.aliasLookup.has(token)) {
      addSiteId(commercialLookups.aliasLookup.get(token).screenCode);
      return;
    }

    if (commercialLookups.screenNameLookup.has(token)) {
      addSiteId(commercialLookups.screenNameLookup.get(token).screenCode);
    }
  });

  return matchedSiteIds;
}

function resolveSiteWiseSubscriptionFee(record) {
  if (!record) {
    return "";
  }

  const pricingMethod = normalizeValue(record?.pricingMethod).toLowerCase();

  // Common pricing is handled by the saved site-wise Complex allocation.
  if (pricingMethod === "common") {
    return "";
  }

  // Allocation Model is screen-wise.
  // Device Count controls OTF Amount only.
  // Plan Mode controls Plan Fee, and that Plan Fee is the site's
  // Subscription Fee used by Stage 3 Billing.
  if (pricingMethod === "allocation model") {
    const allocationRows = Array.isArray(record?.allocationRows)
      ? record.allocationRows
      : [];
    const currentSiteId = getCommercialLookupToken(record?.screenCode);
    const currentScreenName = getCommercialLookupToken(record?.screenName);

    const matchedAllocationRow = allocationRows.find((row) => {
      const rowScreenCode = getCommercialLookupToken(row?.screenCode);
      const rowScreenName = getCommercialLookupToken(row?.screenName);

      return (
        (currentSiteId && rowScreenCode === currentSiteId) ||
        (currentScreenName && rowScreenName === currentScreenName)
      );
    });

    if (matchedAllocationRow) {
      const savedPlanFee = normalizeValue(matchedAllocationRow?.planFee);

      if (savedPlanFee) {
        return savedPlanFee;
      }

      // Backward/fallback support if an older saved Allocation row has
      // Plan Mode but not its calculated Plan Fee snapshot.
      const planMode = normalizeValue(matchedAllocationRow?.planMode);
      const settings = record?.allocationSettingsSnapshot || {};

      if (planMode === "More Than 3") {
        return normalizeValue(settings?.moreThan3PlanFee);
      }

      if (planMode === "Up to 3") {
        return normalizeValue(settings?.upTo3PlanFee);
      }
    }

    return normalizeValue(record?.subscriptionFee);
  }

  const pricingGroups = Array.isArray(record?.pricingGroups)
    ? record.pricingGroups
    : [];

  if (pricingGroups.length === 0) {
    return normalizeValue(record?.subscriptionFee);
  }

  const commercialLookups = getStage1CommercialSiteLookups(record);
  const currentSiteId = getCommercialLookupToken(record?.screenCode);
  const currentScreenName = getCommercialLookupToken(record?.screenName);

  const matchedPricingGroup = pricingGroups.find((group) => {
    const candidateSiteIds = resolvePricingGroupSiteIds(
      group,
      record,
      commercialLookups,
    );

    return (
      candidateSiteIds.includes(currentSiteId) ||
      candidateSiteIds.includes(currentScreenName)
    );
  });

  if (matchedPricingGroup) {
    return normalizeValue(matchedPricingGroup.subscriptionFee);
  }

  return normalizeValue(record?.subscriptionFee);
}

function resolveOtfCommercialValues(record) {
  const pricingMethod = normalizeValue(record?.pricingMethod) || "Common";
  const topLevelApplicable = normalizeValue(record?.otfApplicable);
  const topLevelType = normalizeValue(record?.otfType);
  const topLevelAmount = normalizeValue(record?.otfAmount);
  const commercialLookups = getStage1CommercialSiteLookups(record);
  const pricingGroups = Array.isArray(record?.pricingGroups)
    ? record.pricingGroups
    : [];
  const hasSiteWiseStructure = pricingGroups.length > 0;

  if (pricingMethod.toLowerCase() === "common") {
    return {
      pricingMethod: "Common",
      effectiveOtfApplicable: topLevelApplicable || "No",
      effectiveOtfType: topLevelType,
      effectiveOtfAmount: topLevelAmount,
      matchedPricingGroup: null,
      commercialLookups,
    };
  }

  const currentSiteId = getCommercialLookupToken(record?.screenCode);
  const currentScreenName = getCommercialLookupToken(record?.screenName);
  const matchedPricingGroup = pricingGroups.find((group) => {
    const candidateSiteIds = resolvePricingGroupSiteIds(
      group,
      record,
      commercialLookups,
    );

    return candidateSiteIds.includes(currentSiteId) ||
      candidateSiteIds.includes(currentScreenName);
  });

  if (matchedPricingGroup) {
    return {
      pricingMethod,
      effectiveOtfApplicable:
        normalizeValue(matchedPricingGroup.otfApplicable) || topLevelApplicable || "No",
      effectiveOtfType:
        normalizeValue(matchedPricingGroup.otfType) || topLevelType,
      effectiveOtfAmount:
        normalizeValue(matchedPricingGroup.otfAmount) || topLevelAmount,
      matchedPricingGroup,
      commercialLookups,
    };
  }

  if (hasSiteWiseStructure) {
    return {
      pricingMethod,
      effectiveOtfApplicable: "No",
      effectiveOtfType: "",
      effectiveOtfAmount: "",
      matchedPricingGroup: null,
      commercialLookups,
    };
  }

  return {
    pricingMethod,
    effectiveOtfApplicable: topLevelApplicable || "No",
    effectiveOtfType: topLevelType,
    effectiveOtfAmount: topLevelAmount,
    matchedPricingGroup: null,
    commercialLookups,
  };
}

function getOtfCommercialStatus(record) {
  return (
    normalizeValue(record?.currentCommercialStatus) ||
    normalizeValue(record?.commercialStatus) ||
    normalizeValue(record?.billingCommercialStatus) ||
    (isFoCOriginRecord(record) ? "FoC" : "Billable")
  );
}

function buildOtfTransactions(records = []) {
  const commonGroups = new Map();
  const transactions = [];

  records.forEach((record) => {
    if (isFoCOriginRecord(record)) {
      return;
    }

    const resolvedCommercialValues = resolveOtfCommercialValues(record);
    const groupKey = getStage1GroupKey(record);
    const screenCode = normalizeValue(record.screenCode);
    const effectiveOtfApplicable = normalizeValue(
      resolvedCommercialValues.effectiveOtfApplicable,
    );

    if (effectiveOtfApplicable !== "Yes") {
      return;
    }

    const normalizedComplexCode = normalizeValue(record.complexCode);

    // Common pricing is consolidated only for a real Complex.
    // A standalone site can still use Common pricing, but it must remain
    // visible as its own screen-level OTF row.
    if (
      resolvedCommercialValues.pricingMethod.toLowerCase() === "common" &&
      normalizedComplexCode
    ) {
      const commonGroupKey =
        groupKey || `${normalizeValue(record.billingCode)}::${normalizedComplexCode}`;
      const existingGroup = commonGroups.get(commonGroupKey) || [];
      commonGroups.set(commonGroupKey, [...existingGroup, record]);
      return;
    }

    const transactionId = getOtfTransactionId(record, resolvedCommercialValues.pricingMethod, groupKey);

    transactions.push({
      transactionId,
      pricingMethod: resolvedCommercialValues.pricingMethod,
      billingCode: normalizeValue(record.billingCode),
      complexCode: normalizeValue(record.complexCode),
      billingName: normalizeValue(record.billingName),
      screenCode: screenCode || "-",
      commercialStatus: getOtfCommercialStatus(record),
      siteScope: screenCode || "-",
      includedSiteIds: screenCode ? [screenCode] : [],
      otfType: resolvedCommercialValues.effectiveOtfType,
      otfAmount: resolvedCommercialValues.effectiveOtfAmount,
      otfInvoiceNumber: normalizeValue(record.otfInvoiceNumber),
      otfInvoiceDate: normalizeValue(record.otfInvoiceDate),
      otfRicbrNumber: normalizeValue(record.otfRicbrNumber),
      otfRicbrCreatedDate: normalizeValue(record.otfRicbrCreatedDate),
      otfReceivedAmount: normalizeValue(
        record.otfReceivedAmount || record.otfPaymentReceivedAmount,
      ),
      sourceRecordIds: [record.recordId],
      screenName: normalizeValue(record.screenName),
      location: normalizeValue(record.location),
      state: normalizeValue(record.state),
      searchIndex: [
        record.billingCode,
        record.complexCode,
        record.screenCode,
        record.screenName,
        record.location,
        record.state,
        record.otfInvoiceNumber,
        record.otfRicbrNumber,
        record.otfRicbrCreatedDate,
      ]
        .join(" ")
        .toLowerCase(),
    });
  });

  commonGroups.forEach((groupRecords, commonGroupKey) => {
    const representative = groupRecords[0];
    const snapshotRows = Array.isArray(representative?.stage1GroupRows)
      ? representative.stage1GroupRows
      : [];
    const includedSiteIds = [
      ...snapshotRows.map((row) => normalizeValue(row.screenCode)),
      ...groupRecords.map((row) => normalizeValue(row.screenCode)),
    ].filter(Boolean);
    const uniqueIncludedSiteIds = Array.from(new Set(includedSiteIds));
    const transactionId = getOtfTransactionId(
      representative,
      "Common",
      commonGroupKey,
    );

    const commonScreenNames = Array.from(
      new Set(groupRecords.map((row) => normalizeValue(row.screenName)).filter(Boolean)),
    );
    const commonLocations = Array.from(
      new Set(groupRecords.map((row) => normalizeValue(row.location)).filter(Boolean)),
    );
    const commonStates = Array.from(
      new Set(groupRecords.map((row) => normalizeValue(row.state)).filter(Boolean)),
    );

    transactions.push({
      transactionId,
      pricingMethod: "Common",
      billingCode: normalizeValue(representative.billingCode),
      complexCode: normalizeValue(representative.complexCode),
      billingName: normalizeValue(representative.billingName),
      // Common complex pricing is represented as one complex row.
      // Individual Screen Codes remain searchable through includedSiteIds.
      screenCode: "-",
      commercialStatus: getOtfCommercialStatus(representative),
      siteScope:
        uniqueIncludedSiteIds.length > 0
          ? `All Applicable Sites (${uniqueIncludedSiteIds.length})`
          : "All Applicable Sites",
      includedSiteIds: uniqueIncludedSiteIds,
      otfType: normalizeValue(representative.otfType),
      otfAmount: normalizeValue(representative.otfAmount),
      otfInvoiceNumber: normalizeValue(representative.otfInvoiceNumber),
      otfInvoiceDate: normalizeValue(representative.otfInvoiceDate),
      otfRicbrNumber: normalizeValue(representative.otfRicbrNumber),
      otfRicbrCreatedDate: normalizeValue(representative.otfRicbrCreatedDate),
      otfReceivedAmount: normalizeValue(
        representative.otfReceivedAmount ||
          representative.otfPaymentReceivedAmount,
      ),
      sourceRecordIds: groupRecords.map((row) => row.recordId),
      screenName:
        normalizeValue(representative.complexName) ||
        normalizeValue(representative.billingName) ||
        normalizeValue(representative.complexCode) ||
        "-",
      location:
        commonLocations.length <= 1
          ? commonLocations[0] || normalizeValue(representative.location)
          : "Multiple",
      state:
        commonStates.length <= 1
          ? commonStates[0] || normalizeValue(representative.state)
          : "Multiple",
      searchIndex: [
        representative.billingCode,
        representative.complexCode,
        representative.otfInvoiceNumber,
        representative.otfRicbrNumber,
        representative.otfRicbrCreatedDate,
        ...uniqueIncludedSiteIds,
        ...commonScreenNames,
        ...commonLocations,
        ...commonStates,
      ]
        .join(" ")
        .toLowerCase(),
    });
  });

  return transactions.sort((left, right) =>
    left.siteScope.localeCompare(right.siteScope, undefined, { numeric: true }),
  );
}

function buildCombinedPaymentDownloadRows(otfTransactions = [], stage3Records = []) {
  const otfRows = otfTransactions.map((transaction) => ({
    "Transaction Type": "OTF",
    "Billing Code / Customer Code": transaction.billingCode || "-",
    "Complex Code": getDisplayComplexCode(transaction.complexCode),
    "Screen Code / Scope": transaction.siteScope || "-",
    "Amount Before GST": transaction.otfAmount || "-",
    "Invoice Number": transaction.otfInvoiceNumber || "-",
    "RICBR Number": transaction.otfRicbrNumber || "-",
    "RICBR Created Date": transaction.otfRicbrCreatedDate || "-",
    "Billing Period From": "-",
    "Billing Period To": "-",
    "Payment Status": "-",
    "Payment Received Date": "-",
  }));

  const billingRows = stage3Records.map((record) => ({
    "Transaction Type": "Subscription Fee",
    "Billing Code / Customer Code": normalizeValue(record.billingCode) || "-",
    "Complex Code": getDisplayComplexCode(record.complexCode),
    "Screen Code / Scope": normalizeValue(record.screenCode) || "-",
    "Amount Before GST": normalizeValue(record.subscriptionFee) || "-",
    "Invoice Number":
      normalizeValue(record.subscriptionInvoiceNumber) ||
      normalizeValue(record.invoiceNumber) ||
      "-",
    "RICBR Number": "-",
    "RICBR Created Date": "-",
    "Billing Period From": normalizeValue(record.billingPeriodFrom) || "-",
    "Billing Period To": normalizeValue(record.billingPeriodTo) || "-",
    "Payment Status": normalizeValue(record.paymentStatus) || "-",
    "Payment Received Date": normalizeValue(record.paymentReceivedDate) || "-",
  }));

  return [...otfRows, ...billingRows];
}

function getBillingValidationErrors(record) {
  if (!record) {
    return [];
  }

  const errors = [];
  const billingPeriodFrom = normalizeValue(record.billingPeriodFrom);
  const billingPeriodTo = normalizeValue(record.billingPeriodTo);
  const paymentStatus = normalizeValue(record.paymentStatus) || "Not Invoiced";
  const invoiceRequiredStatuses = new Set([
    "Invoiced",
    "Partially Paid",
    "Paid",
    "Overdue",
    "Disputed",
  ]);
  const invalidDateFields = [
    record.billingStartDate,
    record.billingPeriodFrom,
    record.billingPeriodTo,
    record.invoiceDate,
    record.paymentReceivedDate,
  ].some((dateValue) => dateValue && !isValidFourDigitYear(dateValue));

  if (invalidDateFields) {
    errors.push("Year must contain exactly 4 digits.");
  }

  const normalizedSubscriptionType = normalizeValue(
    record.subscriptionType,
  ).toLowerCase();
  const normalizedSubscriptionMode = normalizeValue(
    record.subscriptionMode,
  )
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const isFixedMonthlyBilling =
    normalizedSubscriptionType === "fixed" &&
    ["monthly", "month", ""].includes(normalizedSubscriptionMode);
  const isFutureBillingPeriodMode = [
    "quarterly",
    "quarter",
    "half year",
    "half yearly",
    "semi annual",
    "yearly",
    "annual",
    "one year",
  ].includes(normalizedSubscriptionMode);
  const billingPeriodFromDate = parseLocalDateValue(record.billingPeriodFrom);
  const billingPeriodToDate = parseLocalDateValue(record.billingPeriodTo);
  const today = new Date();
  const isCurrentMonthDate = (dateValue) =>
    Boolean(dateValue) &&
    dateValue.getFullYear() === today.getFullYear() &&
    dateValue.getMonth() === today.getMonth();
  const allowCurrentMonthFixedPeriodFrom =
    isFixedMonthlyBilling && isCurrentMonthDate(billingPeriodFromDate);
  const allowCurrentMonthFixedPeriodTo =
    isFixedMonthlyBilling && isCurrentMonthDate(billingPeriodToDate);
  const allowFutureBillingPeriodTo = isFutureBillingPeriodMode;

  const futureDateFields = [
    ...(!allowCurrentMonthFixedPeriodFrom
      ? [["Billing Period From", record.billingPeriodFrom]]
      : []),
    ...(!allowCurrentMonthFixedPeriodTo && !allowFutureBillingPeriodTo
      ? [["Billing Period To", record.billingPeriodTo]]
      : []),
    ["Invoice Date", record.invoiceDate],
    ["Payment Received Date", record.paymentReceivedDate],
  ];

  const futureDateField = futureDateFields.find(
    ([, dateValue]) => dateValue && isFutureLocalDateValue(dateValue),
  );

  if (futureDateField) {
    errors.push(`${futureDateField[0]} must not be a future date.`);
  }

  if (billingPeriodFrom && billingPeriodTo) {
    const fromDate = parseDate(billingPeriodFrom);
    const toDate = parseDate(billingPeriodTo);

    if (fromDate && toDate && toDate < fromDate) {
      errors.push("Billing Period To must be on or after Billing Period From.");
    }
  }

  if (invoiceRequiredStatuses.has(paymentStatus)) {
    if (!normalizeValue(record.invoiceNumber)) {
      errors.push("Invoice Number is required for the selected payment status.");
    }

    if (!normalizeValue(record.invoiceDate)) {
      errors.push("Invoice Date is required for the selected payment status.");
    }
  }

  if (
    isPaymentReceivedStatus(paymentStatus) &&
    !normalizeValue(record.paymentReceivedDate)
  ) {
    errors.push(
      "Payment Received Date is required when status is Paid or Payment Received.",
    );
  }

  if (paymentStatus === "Disputed" && !normalizeValue(record.billingRemarks)) {
    errors.push("Billing Remarks are required when status is Disputed.");
  }

  return errors;
}

function Billings({
  stage1Records = [],
  stage2Records = [],
  setStage2Records = () => {},
  stage3Records = [],
  canonicalSiteRecords = [],
  setStage3Records = () => {},
  billingRecords = [],
  setBillingRecords = () => {},
  commonComplexAllocations = {},
  setCommonComplexAllocations = () => {},
  onReturnToStage2 = () => false,
  onReturnCommercialCorrectionToStage1 = () => false,
}) {
  const commonScopeMembershipRecords = [...stage1Records, ...stage2Records];
  const billingRecordSiteCandidates = useMemo(
    () => [...stage3Records, ...canonicalSiteRecords],
    [canonicalSiteRecords, stage3Records],
  );
  const [activeWorkspace, setActiveWorkspace] = useState(workspaceModes.BILLING);
  const [activeBillingSection, setActiveBillingSection] = useState(
    billingSections.FIRST_TIME,
  );
  const [searchTerm, setSearchTerm] = useState("");
  const [firstTimeStatusFilter, setFirstTimeStatusFilter] = useState("All");
  const [recurringSearchTerm, setRecurringSearchTerm] = useState("");
  const [billingRecordsSearchTerm, setBillingRecordsSearchTerm] = useState("");
  const [closureSearchTerm, setClosureSearchTerm] = useState("");
  const [otfSearchTerm, setOtfSearchTerm] = useState("");
  const [otherExpensesSearchTerm, setOtherExpensesSearchTerm] = useState("");
  const [firstTimeFromDate, setFirstTimeFromDate] = useState("");
  const [firstTimeToDate, setFirstTimeToDate] = useState("");
  const [priceChangeSearchTerm, setPriceChangeSearchTerm] = useState("");
  const [priceChangeSortOrder, setPriceChangeSortOrder] = useState("desc");
  const [priceChangeSortKey, setPriceChangeSortKey] = useState("");
  const [reallocationOriginalAllocations, setReallocationOriginalAllocations] = useState({});
  const [priceChangeDraft, setPriceChangeDraft] = useState({
    changeType: "Price Change",
    effectiveDate: "",
    remarks: "",
    newFee: "",
    newMode: "",
  });
  const [priceChangeEffectiveDateDisplay, setPriceChangeEffectiveDateDisplay] =
    useState("");
  const [priceChangeMessage, setPriceChangeMessage] = useState("");
  const [allocationValidationPopup, setAllocationValidationPopup] = useState(null);
  const [isApplyingPriceChange, setIsApplyingPriceChange] = useState(false);
  const [allocationModelChangeDraft, setAllocationModelChangeDraft] = useState({
    modelId: "",
    effectiveDate: "",
    remarks: "",
    planSelections: {},
  });
  const [allocationModelValidation, setAllocationModelValidation] = useState(null);
  const [configuredAllocationModels, setConfiguredAllocationModels] = useState([]);
  const [allocationBackendStatus, setAllocationBackendStatus] = useState("");
  const [recurringSortOrder, setRecurringSortOrder] = useState("desc");
  const [recurringSortKey, setRecurringSortKey] = useState("");
  const [billingRecordsFromDate, setBillingRecordsFromDate] = useState("");
  const [billingRecordsToDate, setBillingRecordsToDate] = useState("");
  const [billingRecordsSortOrder, setBillingRecordsSortOrder] = useState("desc");
  const [billingRecordsSortKey, setBillingRecordsSortKey] = useState("");
  const [closureSortOrder, setClosureSortOrder] = useState("desc");
  const [closureSortKey, setClosureSortKey] = useState("");
  const [otfFromDate, setOtfFromDate] = useState("");
  const [otfToDate, setOtfToDate] = useState("");
  const [otfSortOrder, setOtfSortOrder] = useState("desc");
  const [otfSortKey, setOtfSortKey] = useState("");
  const [otfRowsPerPage, setOtfRowsPerPage] = useState(25);
  const [otfPage, setOtfPage] = useState(1);
  const [otherExpensesFromDate, setOtherExpensesFromDate] = useState("");
  const [otherExpensesToDate, setOtherExpensesToDate] = useState("");
  const [otherExpensesSortOrder, setOtherExpensesSortOrder] = useState("desc");
  const [otherExpensesSortKey, setOtherExpensesSortKey] = useState("");
  const [otherExpensesRowsPerPage, setOtherExpensesRowsPerPage] = useState(25);
  const [otherExpensesPage, setOtherExpensesPage] = useState(1);
  const [persistedExpenseSites, setPersistedExpenseSites] = useState([]);
  const [focFromDate, setFocFromDate] = useState("");
  const [focToDate, setFocToDate] = useState("");
  const [focSortOrder, setFocSortOrder] = useState("desc");
  const [focSortKey, setFocSortKey] = useState("");
  const [focRowsPerPage, setFocRowsPerPage] = useState(25);
  const [focPage, setFocPage] = useState(1);
  const [activeOtherExpenseRecordId, setActiveOtherExpenseRecordId] = useState("");
  const [firstTimeRowsPerPage, setFirstTimeRowsPerPage] = useState(25);
  const [firstTimePage, setFirstTimePage] = useState(1);
  const [recurringRowsPerPage, setRecurringRowsPerPage] = useState(25);
  const [recurringPage, setRecurringPage] = useState(1);
  const [billingRecordsRowsPerPage, setBillingRecordsRowsPerPage] = useState(25);
  const [billingRecordsPage, setBillingRecordsPage] = useState(1);
  const [closureRowsPerPage, setClosureRowsPerPage] = useState(25);
  const [closurePage, setClosurePage] = useState(1);
  const [activeClosureRecordId, setActiveClosureRecordId] = useState("");
  const [closureEditorMode, setClosureEditorMode] = useState("reason");
  const [closureDraft, setClosureDraft] = useState({
    closureType: "",
    effectiveDate: "",
    remarks: "",
  });
  const [restoreActiveDraft, setRestoreActiveDraft] = useState({
    activeFrom: "",
    remarks: "",
  });
  const [closureMessage, setClosureMessage] = useState("");
  const [activeRecordId, setActiveRecordId] = useState("");
  const [activeBillingRecordId, setActiveBillingRecordId] = useState("");
  const [billingRecordDraft, setBillingRecordDraft] = useState(null);
  const [billingRecordSaveMessage, setBillingRecordSaveMessage] = useState("");
  const [billingRecordEditMode, setBillingRecordEditMode] = useState(false);
  const [billingRecordFinancialYear, setBillingRecordFinancialYear] = useState("");
  const [billingRecordMonthFilter, setBillingRecordMonthFilter] = useState("All");
  const [billingRecordSelectedEntryKey, setBillingRecordSelectedEntryKey] =
    useState("");
  const [billingRecordMonthDetailMode, setBillingRecordMonthDetailMode] =
    useState("view");
  const [isErpMonthlyFullscreen, setIsErpMonthlyFullscreen] = useState(false);
  const [billingRecordManualPaymentDraft, setBillingRecordManualPaymentDraft] =
    useState(null);
  const [billingRecordManualPaymentMessage, setBillingRecordManualPaymentMessage] =
    useState("");
  const [billingHistoryScreenCode, setBillingHistoryScreenCode] = useState("");
  const [selectedVerificationIds, setSelectedVerificationIds] = useState([]);
  const [activeOtfTransactionId, setActiveOtfTransactionId] = useState("");
  const [otfTransactionEdits, setOtfTransactionEdits] = useState({});
  const [activeAllocationSiteKey, setActiveAllocationSiteKey] = useState("");
  const [allocationDraftFee, setAllocationDraftFee] = useState("");
  const [allocationSaveMessage, setAllocationSaveMessage] = useState("");
  const verificationSelectAllRef = useRef(null);
  const firstTimeBillingRef = useRef(null);
  const billingInformationRef = useRef(null);
  const commonComplexAllocationRef = useRef(null);
  const priceChangeReallocationRef = useRef(null);
  const priceChangeAllocationRef = useRef(null);
  const priceChangeTransactionIdRef = useRef("");
  const priceChangeApplyPendingRef = useRef(false);
  const recurringBillingWorkspaceRef = useRef(null);
  const [pendingBillingNavigationTarget, setPendingBillingNavigationTarget] =
    useState("");
  const [recurringDecision, setRecurringDecision] = useState("Process");
  const [recurringRemarks, setRecurringRemarks] = useState("");
  const [humanErrorCorrectionRecordId, setHumanErrorCorrectionRecordId] =
    useState("");
  const [humanErrorCorrectionFieldsSelected, setHumanErrorCorrectionFieldsSelected] =
    useState([]);
  const [humanErrorCorrectionRemarks, setHumanErrorCorrectionRemarks] =
    useState("");
  const [firstTimeValidationMessage, setFirstTimeValidationMessage] =
    useState("");
  const firstTimeValidationInFlightRef = useRef(false);
  const [firstTimeWaiverDraft, setFirstTimeWaiverDraft] = useState({
    applicable: "No",
    from: "",
    to: "",
    reason: "",
  });
  const [focSearchTerm, setFocSearchTerm] = useState("");
  const [activeFocRecordId, setActiveFocRecordId] = useState("");
  const [isFocBillableEditorOpen, setIsFocBillableEditorOpen] = useState(false);
  const [focBillableDraft, setFocBillableDraft] = useState({
    effectiveFrom: "",
    otfApplicable: "No",
    otfAmount: "",
    subscriptionApplicable: "Yes",
    subscriptionFee: "",
    remarks: "",
  });
  const [focSaveMessage, setFocSaveMessage] = useState("");
  const [canonicalBillingSummary, setCanonicalBillingSummary] = useState(null);
  const [billingSummaryError, setBillingSummaryError] = useState("");
  const [billingSummaryLoading, setBillingSummaryLoading] = useState(true);

  const normalizedSearch = searchTerm.trim().toLowerCase();
  const normalizedPriceChangeSearch = priceChangeSearchTerm.trim().toLowerCase();
  const normalizedRecurringSearch = recurringSearchTerm.trim().toLowerCase();
  const normalizedBillingRecordsSearch = billingRecordsSearchTerm
    .trim()
    .toLowerCase();
  const normalizedClosureSearch = closureSearchTerm.trim().toLowerCase();
  const normalizedOtfSearch = otfSearchTerm.trim().toLowerCase();

  useEffect(() => {
    let cancelled = false;

    async function loadBillingSummary() {
      setBillingSummaryLoading(true);
      try {
        const summary = await getStage3BillingSummary();
        if (cancelled) return;
        setCanonicalBillingSummary(normalizeStage3BillingSummary(summary));
        setBillingSummaryError("");
      } catch (error) {
        if (cancelled) return;
        console.error("Unable to load canonical Stage 3 Billing summary", error);
        setBillingSummaryError(`Billing summary unavailable. ${error.message}`);
      } finally {
        if (!cancelled) setBillingSummaryLoading(false);
      }
    }

    const handleSummaryUpdated = () => loadBillingSummary();
    loadBillingSummary();
    window.addEventListener("billing-summary-updated", handleSummaryUpdated);
    return () => {
      cancelled = true;
      window.removeEventListener("billing-summary-updated", handleSummaryUpdated);
    };
  }, []);

  const canonicalBillingRecordRows = useMemo(() => {
    const records = canonicalBillingSummary?.billingRecords?.records;
    if (!Array.isArray(records)) return null;

    return records.map((summaryRecord) => {
      const backendSiteId = normalizeValue(summaryRecord?.backendSiteId);
      const screenCode = normalizeValue(summaryRecord?.screenCode).toUpperCase();
      const stage3Record = resolveBillingRecordSite(
        summaryRecord,
        billingRecordSiteCandidates,
      );
      const persisted = stage3Record?.firstTimeBillingRecord;
      const localRecord = billingRecords.find(
        (record) =>
          normalizeValue(record?.billingRecordId) ===
          normalizeValue(summaryRecord?.firstTimeBillingRecordId),
      );

      return {
        ...(stage3Record || {}),
        ...(persisted && typeof persisted === "object" ? persisted : {}),
        ...(localRecord || {}),
        ...summaryRecord,
        backendSiteId:
          backendSiteId ||
          normalizeValue(stage3Record?.backendSiteId) ||
          normalizeValue(persisted?.backendSiteId) ||
          normalizeValue(localRecord?.backendSiteId),
        billingCode:
          normalizeValue(summaryRecord?.billingCode) ||
          normalizeValue(stage3Record?.billingCode) ||
          normalizeValue(localRecord?.billingCode),
        complexCode:
          normalizeValue(summaryRecord?.complexCode) ||
          normalizeValue(stage3Record?.complexCode) ||
          normalizeValue(localRecord?.complexCode),
        screenCode: summaryRecord.screenCode || stage3Record?.screenCode || screenCode,
        screenName:
          normalizeValue(summaryRecord?.screenName) ||
          normalizeValue(stage3Record?.screenName) ||
          normalizeValue(localRecord?.screenName),
        location:
          normalizeValue(summaryRecord?.location) ||
          normalizeValue(stage3Record?.location) ||
          normalizeValue(localRecord?.location),
        submissionStatus:
          normalizeValue(localRecord?.submissionStatus) ||
          normalizeValue(summaryRecord?.billingVerificationStatus),
        billingRecordId:
          normalizeValue(localRecord?.billingRecordId) ||
          normalizeValue(summaryRecord?.firstTimeBillingRecordId),
      };
    });
  }, [billingRecordSiteCandidates, billingRecords, canonicalBillingSummary]);

  useEffect(() => {
    let cancelled = false;

    async function loadAllocationBackendModels() {
      try {
        const migration = await ensureAllocationBackendMigration();
        const records = Array.isArray(migration?.records)
          ? migration.records
          : await billingApiRequest("/settings/allocations");
        if (cancelled) return;

        setConfiguredAllocationModels(
          (Array.isArray(records) ? records : [])
            .map((model, index) =>
              normalizeAllocationModel(model, `Allocation Model ${index + 1}`),
            )
            .filter(Boolean),
        );
        setAllocationBackendStatus("");
      } catch (error) {
        if (!cancelled) {
          console.error("Unable to load Allocation models from backend", error);
          setConfiguredAllocationModels([]);
          setAllocationBackendStatus(
            `Allocation backend unavailable. ${error.message}`,
          );
        }
      }
    }

    const handleAllocationSettingsUpdated = () => {
      loadAllocationBackendModels();
    };

    loadAllocationBackendModels();
    window.addEventListener(
      "billing-allocation-settings-updated",
      handleAllocationSettingsUpdated,
    );

    return () => {
      cancelled = true;
      window.removeEventListener(
        "billing-allocation-settings-updated",
        handleAllocationSettingsUpdated,
      );
    };
  }, []);

  const combinedCommercialRecords = useMemo(() => {
    const nextRecords = new Map();

    [...stage2Records, ...stage3Records].forEach((record) => {
      if (!record?.recordId || nextRecords.has(record.recordId)) {
        return;
      }

      nextRecords.set(record.recordId, record);
    });

    return Array.from(nextRecords.values());
  }, [stage2Records, stage3Records]);

  const rawOtfTransactions = useMemo(
    () => buildOtfTransactions(combinedCommercialRecords),
    [combinedCommercialRecords],
  );

  const otfTransactions = useMemo(
    () =>
      rawOtfTransactions.map((transaction) => ({
        ...transaction,
        ...otfTransactionEdits[transaction.transactionId],
      })),
    [rawOtfTransactions, otfTransactionEdits],
  );

  const verificationRows = useMemo(
    () =>
      stage3Records
        .filter(
          (record) =>
            !isFoCOriginRecord(record) &&
            isSubscriptionApplicableRecord(record),
        )
        .map((record) => {
        const billingTreatment = getDerivedBillingTreatment(record);

        return {
          ...record,
          billingTreatment:
            normalizeValue(record.billingTreatment) || billingTreatment || "",
          billingVerificationStatus:
            normalizeValue(record.correctionType) === "Human Error" &&
            normalizeValue(record.correctionRequestedAt) &&
            !normalizeValue(record.correctionReverifiedAt)
              ? "First Billing Pending Approval"
              : normalizeValue(record.billingVerificationStatus) ||
                "First Billing Pending Approval",
        };
      }),
    [stage3Records],
  );

  const filteredOtfTransactions = useMemo(() => {
    const from = otfFromDate ? parseLocalDateValue(otfFromDate) : null;
    const to = otfToDate ? parseLocalDateValue(otfToDate) : null;

    return otfTransactions
      .filter((transaction) => {
        if (normalizedOtfSearch) {
          const searchableValues = [
            // Billing ID/Code stays searchable even though it is not displayed.
            transaction.billingCode,
            transaction.billingName,
            transaction.complexCode,
            transaction.screenCode,
            transaction.screenName,
            transaction.location,
            transaction.state,
            transaction.commercialStatus,
            transaction.otfInvoiceNumber,
            transaction.otfRicbrNumber,
            transaction.searchIndex,
            ...(transaction.includedSiteIds || []),
          ];
          const matches = searchableValues
            .map((value) => normalizeValue(value).toLowerCase())
            .some((value) => value.includes(normalizedOtfSearch));
          if (!matches) return false;
        }

        if (from || to) {
          const value =
            normalizeValue(transaction.otfInvoiceDate) ||
            normalizeValue(transaction.otfRicbrDate) ||
            normalizeValue(transaction.billingStartDate);
          const date = parseLocalDateValue(value);
          if (!date) return false;
          if (from && date < from) return false;
          if (to && date > to) return false;
        }
        return true;
      })
      .sort((left, right) => {
        const valueFor = (record) => {
          if (!otfSortKey) return normalizeValue(record.otfInvoiceDate) || normalizeValue(record.otfRicbrDate) || normalizeValue(record.billingStartDate);
          return record?.[otfSortKey];
        };
        const diff = compareColumnValues(valueFor(left), valueFor(right));
        return otfSortOrder === "asc" ? diff : -diff;
      });
  }, [normalizedOtfSearch, otfTransactions, otfFromDate, otfToDate, otfSortOrder, otfSortKey]);

  const otfPageCount = Math.max(1, Math.ceil(filteredOtfTransactions.length / otfRowsPerPage));
  const pagedOtfTransactions = useMemo(() => {
    const startIndex = (otfPage - 1) * otfRowsPerPage;
    return filteredOtfTransactions.slice(startIndex, startIndex + otfRowsPerPage);
  }, [filteredOtfTransactions, otfPage, otfRowsPerPage]);

  useEffect(() => {
    setOtfPage(1);
  }, [otfSearchTerm, otfFromDate, otfToDate, otfSortKey, otfSortOrder, otfRowsPerPage]);

  useEffect(() => {
    setOtfPage((currentPage) => Math.min(currentPage, otfPageCount));
  }, [otfPageCount]);

  const filteredVerificationRows = useMemo(
    () =>
      verificationRows.filter((record) => {
        const recordText = [
          record.billingCode,
          record.complexCode,
          record.screenCode,
          record.screenName,
          record.billingName,
          record.location,
          record.correctionType,
        ]
          .join(" ")
          .toLowerCase();

        return normalizedSearch === "" || recordText.includes(normalizedSearch);
      }),
    [verificationRows, normalizedSearch],
  );

  const firstTimeVisibleRows = useMemo(
    () =>
      filteredVerificationRows.filter((record) => {
        const from = firstTimeFromDate ? parseLocalDateValue(firstTimeFromDate) : null;
        const to = firstTimeToDate ? parseLocalDateValue(firstTimeToDate) : null;
        if (from || to) {
          const billingDate = parseLocalDateValue(normalizeValue(record.billingStartDate));
          if (!billingDate) return false;
          if (from && billingDate < from) return false;
          if (to && billingDate > to) return false;
        }
        const isPending = isPendingFirstTimeBillingRecord(record);
        const persistedStatus = getPersistedFirstTimeBillingStatus(record);
        const isValidated =
          persistedStatus === "Validated" &&
          isPending &&
          isFirstTimeBillingSiteValidatedForCurrentCycle(record);
        const isCompleted =
          (persistedStatus === "Completed" &&
            isFirstBillingSubmittedInCurrentMonth(record, billingRecords)) ||
          (!isPending &&
            isFirstBillingSubmittedInCurrentMonth(record, billingRecords));

        if (!isPending && !isCompleted) {
          return false;
        }

        if (firstTimeStatusFilter === "Pending") {
          return isPending && !isValidated;
        }

        if (firstTimeStatusFilter === "Validated") {
          return isValidated;
        }

        if (firstTimeStatusFilter === "Completed") {
          return isCompleted;
        }

        return true;
      }),
    [filteredVerificationRows, billingRecords, firstTimeStatusFilter, firstTimeFromDate, firstTimeToDate],
  );

  const selectableVerificationRows = useMemo(
    () =>
      firstTimeVisibleRows.filter((record) =>
        isFirstTimeBillingSiteReadyForSubmission(
          record,
          stage3Records,
          commonComplexAllocations,
        ),
      ),
    [
      firstTimeVisibleRows,
      stage3Records,
      commonComplexAllocations,
    ],
  );

  const firstTimeBillingDownloadRows = useMemo(
    () =>
      firstTimeVisibleRows.filter(
        (record) =>
          !isPendingFirstTimeBillingRecord(record) &&
          isFirstBillingSubmittedInCurrentMonth(record, billingRecords),
      ),
    [firstTimeVisibleRows, billingRecords],
  );

  const firstTimePageCount = Math.max(
    1,
    Math.ceil(firstTimeVisibleRows.length / firstTimeRowsPerPage),
  );
  const pagedVerificationRows = useMemo(() => {
    const startIndex = (firstTimePage - 1) * firstTimeRowsPerPage;
    return firstTimeVisibleRows.slice(
      startIndex,
      startIndex + firstTimeRowsPerPage,
    );
  }, [firstTimeVisibleRows, firstTimePage, firstTimeRowsPerPage]);

  const siteRecurringRows = useMemo(
    () =>
      verificationRows
        .filter(
          (record) =>
            !isPendingFirstTimeBillingRecord(record) &&
            isFirstBillingBeforeCurrentMonth(record, billingRecords),
        )
        .map((record) => {
          const recurringCycle = getRecurringBillingCycle(
            record,
            billingRecords,
            stage3Records,
            commonComplexAllocations,
          );

          const reason = normalizeValue(record.billingLifecycleReason);
          const billingStatus =
            normalizeValue(record.billingLifecycleStatus) ||
            (reason ? "Inactive" : "Active");

          if (billingStatus === "Active" && !reason) {
            return {
              ...record,
              recurringCycle,
            };
          }

          const cutoffValue = normalizeValue(
            reason === "Billing Paused"
              ? record.pauseFromDate || record.closureEffectiveDate
              : record.closureEffectiveDate ||
                  record.inactiveEffectiveDate ||
                  (record.inactiveSince
                    ? String(record.inactiveSince).slice(0, 10)
                    : ""),
          );
          const cutoffDate = parseLocalDateValue(cutoffValue);
          const periodFromDate = parseLocalDateValue(recurringCycle.periodFrom);
          const normalPeriodToDate = parseLocalDateValue(recurringCycle.periodTo);

          if (!cutoffDate || !periodFromDate || !normalPeriodToDate) {
            return null;
          }

          const finalBillThrough = new Date(cutoffDate);
          finalBillThrough.setDate(finalBillThrough.getDate() - 1);

          if (finalBillThrough < periodFromDate) {
            return null;
          }

          const adjustedPeriodTo =
            finalBillThrough < normalPeriodToDate
              ? finalBillThrough
              : normalPeriodToDate;

          return {
            ...record,
            recurringCycle: {
              ...recurringCycle,
              periodTo: formatLocalDateValue(adjustedPeriodTo),
              eligibilityStatus: "Final Billing Before Inactive",
              finalBillingCutoff: cutoffValue,
            },
          };
        })
        .filter(
          (record) =>
            record &&
            record.recurringCycle?.currentMonthDue === true &&
            Boolean(record.recurringCycle?.periodFrom) &&
            Boolean(record.recurringCycle?.periodTo),
        ),
    [
      verificationRows,
      billingRecords,
      stage3Records,
      commonComplexAllocations,
    ],
  );


  const recurringRows = useMemo(() => {
    if (Array.isArray(canonicalBillingSummary?.recurringBilling?.records)) {
      return canonicalBillingSummary.recurringBilling.records;
    }

    const nextRows = [];
    const combinedGroups = new Map();

    siteRecurringRows.forEach((record) => {
      if (
        !isCommonComplexBillingRecord(record) ||
        Boolean(
          getCommonComplexBillingValidationError(
            record,
            stage3Records,
            commonComplexAllocations,
          ),
        )
      ) {
        nextRows.push(record);
        return;
      }

      const groupKey = getCommonComplexBillingGroupKey(record);
      if (!groupKey) {
        nextRows.push(record);
        return;
      }

      const currentRows = combinedGroups.get(groupKey) || [];
      combinedGroups.set(groupKey, [...currentRows, record]);
    });

    combinedGroups.forEach((groupRows, groupKey) => {
      if (groupRows.length === 0) return;

      const representative = groupRows[0];
      const alignment = getCommonComplexBillingAlignment(
        representative,
        stage3Records,
        commonComplexAllocations,
      );
      const commonSubscriptionFee = resolveCommonComplexSubscriptionFee(
        representative,
        stage3Records,
        commonComplexAllocations,
      );
      const billingAmounts = groupRows
        .map((row) => parseAmountValue(row.recurringCycle?.billingSubscriptionFee))
        .filter((value) => value !== null);
      const combinedBillingAmount = billingAmounts.reduce(
        (sum, value) => sum + value,
        0,
      );
      const fromDates = groupRows
        .map((row) => parseLocalDateValue(row.recurringCycle?.periodFrom))
        .filter(Boolean);
      const toDates = groupRows
        .map((row) => parseLocalDateValue(row.recurringCycle?.periodTo))
        .filter(Boolean);
      const combinedPeriodFrom = fromDates.length
        ? formatLocalDateValue(new Date(Math.min(...fromDates.map((date) => date.getTime()))))
        : "";
      const combinedPeriodTo = toDates.length
        ? formatLocalDateValue(new Date(Math.max(...toDates.map((date) => date.getTime()))))
        : "";
      const locations = Array.from(
        new Set(groupRows.map((row) => normalizeValue(row.location)).filter(Boolean)),
      );
      const blockedChild = groupRows.find(
        (row) => row.recurringCycle?.priceChangeBlocked === true,
      );
      const hasOverdueChild = groupRows.some(
        (row) => row.recurringCycle?.overdue === true,
      );

      nextRows.push({
        ...representative,
        recordId: `COMMON-RECURRING::${groupKey}`,
        isCombinedCommonBilling: true,
        combinedGroupKey: groupKey,
        combinedChildRecords: groupRows,
        screenCode: "Combined",
        screenName:
          normalizeValue(representative.complexName) ||
          normalizeValue(representative.billingName) ||
          normalizeValue(representative.complexCode) ||
          "Complex",
        location: locations.length <= 1 ? locations[0] || "" : "Multiple",
        recurringCycle: {
          ...representative.recurringCycle,
          periodFrom: combinedPeriodFrom,
          periodTo: combinedPeriodTo,
          billingMonth: representative.recurringCycle?.billingMonth || "",
          actualSubscriptionFee: normalizeValue(commonSubscriptionFee),
          billingSubscriptionFee: formatAmountValue(combinedBillingAmount),
          eligibilityStatus: blockedChild
            ? "Price Change & Reallocation Pending"
            : hasOverdueChild
              ? "Overdue / Unbilled"
              : "Combined Complex Billing",
          commonBillingStartDate: alignment.commonBillingStartDate,
          eligible: !blockedChild,
          currentMonthDue: true,
          billingWorkStatus: "Unbilled",
          overdue: hasOverdueChild,
          priceChangeBlocked: Boolean(blockedChild),
          priceChangeBlockScope: blockedChild?.recurringCycle?.priceChangeBlockScope || "",
          priceChangeBlockLabel: blockedChild?.recurringCycle?.priceChangeBlockLabel || "",
        },
      });
    });

    return nextRows;
  }, [
    canonicalBillingSummary,
    siteRecurringRows,
    stage3Records,
    commonComplexAllocations,
  ]);

  function toggleColumnSort(activeKey, setKey, direction, setDirection, key) {
    if (activeKey === key) setDirection(direction === "asc" ? "desc" : "asc");
    else { setKey(key); setDirection("asc"); }
  }

  const filteredRecurringRows = useMemo(() => {
    const statusQuery = normalizedRecurringSearch;
    const billedTerms = ["billed", "processed", "completed"];
    const unbilledTerms = ["unbilled", "pending", "overdue", "billing pending"];

    const filtered = recurringRows.filter((record) => {
      const statusText = normalizeValue(
        record.recurringCycle?.billingWorkStatus ||
          record.recurringCycle?.eligibilityStatus,
      ).toLowerCase();

      if (statusQuery && billedTerms.includes(statusQuery)) {
        return statusText.includes("billed") && !statusText.includes("unbilled");
      }

      if (statusQuery && unbilledTerms.includes(statusQuery)) {
        return (
          statusText.includes("unbilled") ||
          statusText.includes("pending") ||
          record.recurringCycle?.overdue === true
        );
      }

      if (!statusQuery) return true;

      return [
        record.billingCode,
        record.complexCode,
        record.screenCode,
        record.screenName,
        record.billingName,
        record.location,
        record.state,
        record.subscriptionType,
        record.subscriptionMode,
        record.recurringCycle?.billingMonth,
        record.recurringCycle?.periodFrom,
        record.recurringCycle?.periodTo,
        record.recurringCycle?.eligibilityStatus,
        record.recurringCycle?.billingWorkStatus,
        record.billingVerificationStatus,
      ]
        .map((value) => normalizeValue(value).toLowerCase())
        .some((value) => value.includes(statusQuery));
    });

    return [...filtered].sort((left, right) => {
      // Default: overdue/unbilled work first. User-selected column sort still wins.
      if (!recurringSortKey) {
        const leftPriority = left.recurringCycle?.overdue === true ? 0 : 1;
        const rightPriority = right.recurringCycle?.overdue === true ? 0 : 1;
        if (leftPriority !== rightPriority) return leftPriority - rightPriority;

        const leftDate = parseLocalDateValue(left.recurringCycle?.periodFrom);
        const rightDate = parseLocalDateValue(right.recurringCycle?.periodFrom);
        return (leftDate?.getTime() || 0) - (rightDate?.getTime() || 0);
      }

      const valueFor = (record) =>
        recurringSortKey === "billingDate"
          ? record.recurringCycle?.periodFrom
          : record?.[recurringSortKey];
      const diff = compareColumnValues(valueFor(left), valueFor(right));
      return recurringSortOrder === "asc" ? diff : -diff;
    });
  }, [normalizedRecurringSearch, recurringRows, recurringSortOrder, recurringSortKey]);

  const recurringBillingDownloadRows = useMemo(
    () => filteredRecurringRows,
    [filteredRecurringRows],
  );

  const activeRecurringRecord = useMemo(
    () =>
      recurringRows.find((record) => record.recordId === activeRecordId) || null,
    [recurringRows, activeRecordId],
  );

  const recurringPageCount = Math.max(
    1,
    Math.ceil(filteredRecurringRows.length / recurringRowsPerPage),
  );
  const pagedRecurringRows = useMemo(() => {
    const startIndex = (recurringPage - 1) * recurringRowsPerPage;
    return filteredRecurringRows.slice(
      startIndex,
      startIndex + recurringRowsPerPage,
    );
  }, [filteredRecurringRows, recurringPage, recurringRowsPerPage]);

  const activeOtfTransaction = useMemo(
    () =>
      otfTransactions.find(
        (transaction) => transaction.transactionId === activeOtfTransactionId,
      ) || null,
    [activeOtfTransactionId, otfTransactions],
  );

  const activeRecord = useMemo(
    () =>
      stage3Records.find((record) => record.recordId === activeRecordId) || null,
    [activeRecordId, stage3Records],
  );

  const humanErrorCorrectionBillingRecord = useMemo(
    () =>
      billingRecords.find(
        (record) => record.billingRecordId === humanErrorCorrectionRecordId,
      ) || null,
    [humanErrorCorrectionRecordId, billingRecords],
  );

  const humanErrorCorrectionRecord = useMemo(() => {
    if (!humanErrorCorrectionBillingRecord) {
      return null;
    }

    const sourceRecordId = normalizeValue(
      humanErrorCorrectionBillingRecord.sourceRecordId,
    );
    const screenCode = normalizeValue(humanErrorCorrectionBillingRecord.screenCode);

    return (
      stage3Records.find(
        (record) =>
          (sourceRecordId && normalizeValue(record.recordId) === sourceRecordId) ||
          (screenCode && normalizeValue(record.screenCode) === screenCode),
      ) || null
    );
  }, [humanErrorCorrectionBillingRecord, stage3Records]);

  useEffect(() => {
    if (
      !pendingBillingNavigationTarget ||
      (!activeRecord && !activeRecurringRecord)
    ) {
      return;
    }

    window.requestAnimationFrame(() => {
      if (pendingBillingNavigationTarget === "recurring-workspace") {
        guidedScrollToElement(recurringBillingWorkspaceRef.current);
      } else if (pendingBillingNavigationTarget === "common-allocation") {
        guidedScrollToElement(commonComplexAllocationRef.current);
      } else {
        guidedScrollToElement(billingInformationRef.current);
      }

      setPendingBillingNavigationTarget("");
    });
  }, [activeRecord, activeRecurringRecord, pendingBillingNavigationTarget]);

  useEffect(() => {
    setActiveWorkspace(workspaceModes.BILLING);
    setActiveBillingSection(billingSections.FIRST_TIME);

    window.requestAnimationFrame(() => {
      guidedScrollToElement(firstTimeBillingRef.current);
    });
  }, []);

  function hasFirstBillingStarted(record) {
    const billingStartDate = parseLocalDateValue(record?.billingStartDate);
    if (!billingStartDate) return false;

    const today = new Date();
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const billingStart = new Date(
      billingStartDate.getFullYear(),
      billingStartDate.getMonth(),
      billingStartDate.getDate(),
    );

    return billingStart <= todayStart;
  }


  function isPriceChangeLifecycleActive(record) {
    if (!record) return false;

    const lifecycleStatus = normalizeValue(record?.billingLifecycleStatus).toLowerCase();
    const lifecycleReason = normalizeValue(
      record?.billingLifecycleReason ||
        record?.billingStatusReason ||
        record?.currentBillingStatus,
    ).toLowerCase();

    if (lifecycleStatus === "inactive") return false;

    return ![
      "billing paused",
      "site inactive",
      "site removed",
      "contract closed",
      "inactive",
      "closed",
    ].includes(lifecycleReason);
  }

  function getStage3RecordForPriceChangeBillingRecord(billingRecord) {
    return resolveBillingRecordSite(billingRecord, billingRecordSiteCandidates);
  }

  function handlePriceChangeTypeSelection(nextChangeType) {
    if (!activeRecord) return;

    const currentMode = normalizeSubscriptionModeLabel(activeRecord.subscriptionMode);
    const defaultAlternative = getDefaultAlternativeSubscriptionMode(currentMode);

    setPriceChangeMessage("");
    setPriceChangeDraft((current) => ({
      ...current,
      changeType: nextChangeType,
      newMode:
        nextChangeType === "Price Change"
          ? currentMode
          : normalizeSubscriptionModeLabel(current.newMode) === currentMode
            ? defaultAlternative
            : normalizeValue(current.newMode) || defaultAlternative,
    }));
  }

  function handleNewSubscriptionModeSelection(nextMode) {
    if (!activeRecord) return;

    const currentMode = normalizeSubscriptionModeLabel(activeRecord.subscriptionMode);
    const normalizedNextMode = normalizeSubscriptionModeLabel(nextMode);

    if (normalizedNextMode === currentMode) {
      const message =
        "New Subscription Mode cannot be the same as Current Subscription Mode.";
      window.alert(message);
      setPriceChangeMessage(message);
      return;
    }

    setPriceChangeMessage("");
    setPriceChangeDraft((current) => ({
      ...current,
      newMode: normalizedNextMode,
    }));
  }

  function handleNewSubscriptionFeeBlur() {
    if (!activeRecord) return;

    const changeType = normalizeValue(priceChangeDraft.changeType) || "Price Change";
    if (!["Price Change", "Price + Mode Change"].includes(changeType)) return;

    const currentFee = getPriceChangeCurrentFee(activeRecord);
    const newFee = parseAmountValue(priceChangeDraft.newFee);

    if (
      currentFee !== null &&
      newFee !== null &&
      Math.abs(currentFee - newFee) < 0.005
    ) {
      const message =
        "New Subscription Fee cannot be the same as Current Subscription Fee.";
      window.alert(message);
      setPriceChangeMessage(message);
    }
  }

  function handlePriceChangeEffectiveDateInput(event) {
    const nextDisplay = maskPriceChangeDateInput(event.target.value);
    const canonicalValue = getCanonicalPriceChangeDate(nextDisplay);
    const minimumValue = getMinimumPriceChangeEffectiveDate(activeRecord);

    setPriceChangeEffectiveDateDisplay(nextDisplay);
    setPriceChangeDraft((current) => ({
      ...current,
      effectiveDate: canonicalValue,
    }));

    if (!canonicalValue) {
      setPriceChangeMessage("");
      return;
    }

    if (minimumValue && canonicalValue < minimumValue) {
      setPriceChangeMessage(
        `Price Change Effective Date must be on or after ${formatBillingDateDisplay(minimumValue)} for this billing mode.`,
      );
      return;
    }

    setPriceChangeMessage("");
  }

  function handlePriceChangeEffectiveDateBlur() {
    if (!priceChangeEffectiveDateDisplay) return;

    if (!getCanonicalPriceChangeDate(priceChangeEffectiveDateDisplay)) {
      setPriceChangeMessage(
        "Please enter a valid Effective Date using a 4-digit year.",
      );
    }
  }

  function getAllocationModelAffectedRecords(record) {
    if (!isAllocationModelBillingRecord(record)) {
      return [record].filter(
        (candidate) => candidate && isPriceChangeLifecycleActive(candidate),
      );
    }

    const groupKey = getAllocationModelBillingGroupKey(record);
    return stage3Records.filter(
      (candidate) =>
        isAllocationModelBillingRecord(candidate) &&
        getAllocationModelBillingGroupKey(candidate) === groupKey &&
        isPriceChangeLifecycleActive(candidate),
    );
  }

  function getAllocationModelCurrentSnapshot(record) {
    return normalizeAllocationModel(
      record?.allocationSettingsSnapshot,
      "Current Allocation Model",
    );
  }

  function getAllocationModelRowForRecord(record) {
    const screenCode = normalizeValue(record?.screenCode).toUpperCase();
    const rows = Array.isArray(record?.allocationRows) ? record.allocationRows : [];
    return (
      rows.find(
        (row) =>
          normalizeValue(row?.screenCode).toUpperCase() === screenCode,
      ) || null
    );
  }

  function getSelectedAllocationModel() {
    const models = activeRecord
      ? getConfiguredAllocationModelsForRecord(
          activeRecord,
          configuredAllocationModels,
        )
      : configuredAllocationModels;
    return (
      models.find(
        (model) =>
          normalizeValue(model.modelId) ===
          normalizeValue(allocationModelChangeDraft.modelId),
      ) || null
    );
  }

  function buildStage3BackendSnapshot(record) {
    return {
      billingLifecycleStatus:
        normalizeValue(record?.billingLifecycleStatus) || "Active",
      billingLifecycleReason: normalizeValue(record?.billingLifecycleReason),
      billingStatusReason: normalizeValue(record?.billingStatusReason),
      currentBillingStatus: normalizeValue(record?.currentBillingStatus),
      closureEffectiveDate: normalizeValue(record?.closureEffectiveDate),
      pauseFromDate: normalizeValue(record?.pauseFromDate),
      billingActiveFromDate: normalizeValue(record?.billingActiveFromDate),
      billingStartDate: normalizeValue(record?.billingStartDate),
      pricingMethod: normalizeValue(record?.pricingMethod),
      subscriptionFee: normalizeValue(record?.subscriptionFee),
      allocationChainNetwork: getAllocationChainName(record),
      allocationSettingsSnapshot:
        record?.allocationSettingsSnapshot &&
        typeof record.allocationSettingsSnapshot === "object"
          ? record.allocationSettingsSnapshot
          : null,
      allocationRows: Array.isArray(record?.allocationRows)
        ? record.allocationRows.map((row) => ({ ...row }))
        : [],
      allocationModelEffectiveDate: normalizeValue(
        record?.allocationModelEffectiveDate,
      ),
      allocationModelChangeStatus: normalizeValue(
        record?.allocationModelChangeStatus,
      ),
      allocationModelHistory: Array.isArray(record?.allocationModelHistory)
        ? record.allocationModelHistory
        : [],
    };
  }

  async function syncStage3RecordsForAllocationValidation(records) {
    const syncable = (Array.isArray(records) ? records : []).filter(
      (record) => normalizeValue(record?.backendSiteId),
    );

    await Promise.all(
      syncable.map((record) =>
        billingApiRequest(`/sites/${record.backendSiteId}`, {
          method: "PATCH",
          body: JSON.stringify({
            stage3Data: buildStage3BackendSnapshot(record),
          }),
        }),
      ),
    );
  }

  function getAllocationModelDraftSignature() {
    const sortedPlanSelections = Object.entries(
      allocationModelChangeDraft.planSelections || {},
    )
      .map(([recordId, plan]) => [recordId, normalizeValue(plan)])
      .sort(([leftId], [rightId]) => leftId.localeCompare(rightId));

    return JSON.stringify({
      modelId: normalizeValue(allocationModelChangeDraft.modelId),
      effectiveDate: normalizeValue(allocationModelChangeDraft.effectiveDate),
      remarks: normalizeValue(allocationModelChangeDraft.remarks),
      planSelections: sortedPlanSelections,
    });
  }

  function isAllocationModelValidationCurrent() {
    return Boolean(
      allocationModelValidation?.signature &&
      allocationModelValidation.signature === getAllocationModelDraftSignature(),
    );
  }

  async function handleValidateAllocationModelChange() {
    if (!activeRecord || !isAllocationModelBillingRecord(activeRecord)) return;

    const selectedModel = getSelectedAllocationModel();
    const currentModel = getAllocationModelCurrentSnapshot(activeRecord);
    const currentChainName = getAllocationChainName(activeRecord);
    const effectiveDateValue = normalizeValue(
      allocationModelChangeDraft.effectiveDate,
    );
    const remarks = normalizeValue(allocationModelChangeDraft.remarks);

    if (!selectedModel) {
      setAllocationModelValidation(null);
      setPriceChangeMessage(
        currentChainName
          ? `Select an Active ${currentChainName} Allocation configured in Settings.`
          : "Select a New Allocation Model configured in Settings.",
      );
      return;
    }

    if (
      currentChainName &&
      getAllocationChainName(selectedModel).toLowerCase() !==
        currentChainName.toLowerCase()
    ) {
      setAllocationModelValidation(null);
      setPriceChangeMessage(
        `The selected Allocation belongs to ${getAllocationChainName(selectedModel)}. Select an Allocation configured for ${currentChainName}.`,
      );
      return;
    }

    const currentSummary = formatAllocationModelSummary(currentModel);
    const newSummary = formatAllocationModelSummary(selectedModel);

    if (currentSummary === newSummary) {
      setAllocationModelValidation(null);
      const message =
        "New Allocation Model must be different from the Current Allocation Model.";
      window.alert(message);
      setPriceChangeMessage(message);
      return;
    }

    if (!effectiveDateValue || !isValidDateValue(effectiveDateValue)) {
      setAllocationModelValidation(null);
      setPriceChangeMessage(
        "A valid Allocation Model Effective Date is required.",
      );
      return;
    }

    if (!remarks) {
      setAllocationModelValidation(null);
      setPriceChangeMessage("Remarks / Reason is required.");
      return;
    }

    const affectedRecords = getAllocationModelAffectedRecords(activeRecord);

    if (affectedRecords.length === 0) {
      setAllocationModelValidation(null);
      setPriceChangeMessage(
        "No Active billing screens are eligible for this Allocation change.",
      );
      return;
    }

    const missingPlanRecord = affectedRecords.find((record) => {
      const selectedPlanLabel = normalizeValue(
        allocationModelChangeDraft.planSelections?.[record.recordId],
      );

      if (!selectedPlanLabel) return true;

      return !(selectedModel?.slabs || []).some(
        (slab) => normalizeValue(slab?.label) === selectedPlanLabel,
      );
    });

    if (missingPlanRecord) {
      setAllocationModelValidation(null);
      setPriceChangeMessage(
        `Select a New Plan for ${
          missingPlanRecord.screenCode ||
          missingPlanRecord.screenName ||
          "each eligible screen"
        }.`,
      );
      return;
    }

    const earliestBillingStart = affectedRecords
      .map((record) => parseLocalDateValue(record?.billingStartDate))
      .filter(Boolean)
      .sort((left, right) => left - right)[0];

    const effectiveDate = parseLocalDateValue(effectiveDateValue);

    if (earliestBillingStart && effectiveDate < earliestBillingStart) {
      setAllocationModelValidation(null);
      setPriceChangeMessage(
        "Allocation Model Effective Date cannot be earlier than the first Billing Start Date in the selected group.",
      );
      return;
    }

    const validatedAt = new Date().toISOString();
    const transactionId =
      priceChangeTransactionIdRef.current || crypto.randomUUID();
    const groupKey = getAllocationModelBillingGroupKey(activeRecord);
    const validatedPlanSelections = {};

    affectedRecords.forEach((record) => {
      validatedPlanSelections[record.recordId] = normalizeValue(
        allocationModelChangeDraft.planSelections?.[record.recordId],
      );
    });

    try {
      await syncStage3RecordsForAllocationValidation(affectedRecords);

      const backendValidation = await billingApiRequest(
        "/stage3/allocation-reallocation/validate",
        {
          method: "POST",
          body: JSON.stringify({
            chainName: currentChainName || getAllocationChainName(selectedModel),
            allocationId: selectedModel.modelId,
            effectiveDate: effectiveDateValue,
            remarks,
            changedBy: "Operations",
            items: affectedRecords.map((record) => ({
              screenCode: normalizeValue(record.screenCode),
              newPlanMode: normalizeValue(
                validatedPlanSelections[record.recordId],
              ),
              deviceCount:
                Number(getAllocationModelRowForRecord(record)?.count) ||
                Number(record?.deviceCount) ||
                null,
            })),
          }),
        },
      );

      setAllocationModelValidation({
        signature: getAllocationModelDraftSignature(),
        validationId:
          normalizeValue(backendValidation?.validationId) ||
          normalizeValue(backendValidation?.id),
        validatedAt:
          normalizeValue(backendValidation?.validatedAt) || validatedAt,
        groupKey,
        modelSnapshot: {
          ...selectedModel,
        },
        effectiveDate: effectiveDateValue,
        remarks,
        planSelections: validatedPlanSelections,
        affectedRecordIds: affectedRecords.map((record) => record.recordId),
        backendValidation,
      });

      setPriceChangeMessage(
        "Allocation Model change validated by backend. Review the proposed Allocation, Plans and Fees, then click Save & Submit to make the change effective.",
      );
    } catch (error) {
      setAllocationModelValidation(null);
      setPriceChangeMessage(
        `Allocation Model validation failed. ${error.message}`,
      );
    }
  }

  async function handleSaveAndSubmitAllocationModelChange() {
    if (!activeRecord || !isAllocationModelBillingRecord(activeRecord)) return;

    if (!allocationModelValidation || !isAllocationModelValidationCurrent()) {
      setPriceChangeMessage(
        "The Allocation values have not been validated, or they changed after validation. Run Validate & Apply Allocation Model again before Save & Submit.",
      );
      return;
    }

    const selectedModel = normalizeAllocationModel(
      allocationModelValidation.modelSnapshot,
      "Validated Allocation Model",
    );

    if (!selectedModel) {
      setAllocationModelValidation(null);
      setPriceChangeMessage(
        "The validated Allocation Model is no longer available. Validate the change again.",
      );
      return;
    }

    const groupKey = allocationModelValidation.groupKey;
    const validatedRecordIds = [
      ...(allocationModelValidation.affectedRecordIds || []),
    ].sort();

    const currentEligibleRecords = stage3Records.filter(
      (record) =>
        isAllocationModelBillingRecord(record) &&
        getAllocationModelBillingGroupKey(record) === groupKey &&
        isPriceChangeLifecycleActive(record),
    );
    const currentEligibleIds = currentEligibleRecords
      .map((record) => record.recordId)
      .sort();

    const eligibleSetChanged =
      validatedRecordIds.length !== currentEligibleIds.length ||
      validatedRecordIds.some(
        (recordId, index) => recordId !== currentEligibleIds[index],
      );

    if (eligibleSetChanged) {
      setAllocationModelValidation(null);
      setPriceChangeMessage(
        "The eligible Active billing screens changed after validation. Validate the Allocation Model change again before saving.",
      );
      return;
    }

    const invalidPlanRecord = currentEligibleRecords.find((record) => {
      const planLabel = normalizeValue(
        allocationModelValidation.planSelections?.[record.recordId],
      );

      return !(selectedModel.slabs || []).some(
        (slab) => normalizeValue(slab?.label) === planLabel,
      );
    });

    if (invalidPlanRecord) {
      setAllocationModelValidation(null);
      setPriceChangeMessage(
        "One or more validated Plans are no longer available. Validate the Allocation Model change again.",
      );
      return;
    }

    const validationId = normalizeValue(allocationModelValidation.validationId);
    if (!validationId) {
      setAllocationModelValidation(null);
      setPriceChangeMessage(
        "Backend Validation ID is missing. Run Validate & Apply Allocation Model again before Save & Submit.",
      );
      return;
    }

    let backendSavedChange;
    try {
      backendSavedChange = await billingApiRequest(
        `/stage3/allocation-reallocation/${validationId}/save-submit`,
        {
          method: "POST",
          body: JSON.stringify({ changedBy: "Operations" }),
        },
      );
    } catch (error) {
      setPriceChangeMessage(
        `Allocation Model Save & Submit failed. No frontend Allocation values were changed. ${error.message}`,
      );
      return;
    }

    const backendItemByScreen = new Map(
      (Array.isArray(backendSavedChange?.items)
        ? backendSavedChange.items
        : []
      ).map((item) => [
        normalizeValue(item?.screenCode).toUpperCase(),
        item,
      ]),
    );

    const submittedAt =
      normalizeValue(backendSavedChange?.savedSubmittedAt) ||
      new Date().toISOString();
    const effectiveDateValue = normalizeValue(
      allocationModelValidation.effectiveDate,
    );
    const remarks = normalizeValue(allocationModelValidation.remarks);

    setStage3Records((currentRecords) =>
      currentRecords.map((record) => {
        if (
          !validatedRecordIds.includes(record.recordId) ||
          !isAllocationModelBillingRecord(record) ||
          getAllocationModelBillingGroupKey(record) !== groupKey ||
          !isPriceChangeLifecycleActive(record)
        ) {
          return record;
        }

        const currentAllocationRow = getAllocationModelRowForRecord(record);
        const selectedPlanLabel = normalizeValue(
          allocationModelValidation.planSelections?.[record.recordId],
        );
        const backendItem = backendItemByScreen.get(
          normalizeValue(record?.screenCode).toUpperCase(),
        );
        const backendPlanLabel = normalizeValue(backendItem?.newPlanMode);
        const effectivePlanLabel = backendPlanLabel || selectedPlanLabel;
        const newSlab =
          selectedModel?.slabs?.find(
            (slab) => normalizeValue(slab?.label) === effectivePlanLabel,
          ) || null;
        const nextPlanFee =
          normalizeValue(backendItem?.newPlanFee) ||
          normalizeValue(newSlab?.planFee);

        const existingRows = Array.isArray(record?.allocationRows)
          ? record.allocationRows
          : [];
        let matchedCurrentRow = false;

        const nextAllocationRows = existingRows.map((row) => {
          if (
            normalizeValue(row?.screenCode).toUpperCase() !==
            normalizeValue(record?.screenCode).toUpperCase()
          ) {
            return row;
          }

          matchedCurrentRow = true;

          return {
            ...row,
            planMode: effectivePlanLabel,
            planFee: nextPlanFee,
          };
        });

        if (!matchedCurrentRow && normalizeValue(record?.screenCode)) {
          nextAllocationRows.push({
            screenCode: normalizeValue(record?.screenCode),
            screenName: normalizeValue(record?.screenName),
            count: Number(record?.deviceCount) || 1,
            planMode: effectivePlanLabel,
            planFee: nextPlanFee,
          });
        }

        const previousSnapshot = getAllocationModelCurrentSnapshot(record);
        const historyEntry = {
          historyId: `allocation-model-change-${Date.now()}-${Math.random()
            .toString(36)
            .slice(2, 8)}`,
          previousModel: previousSnapshot,
          newModel: {
            ...selectedModel,
            otfPricePerDevice: normalizeValue(
              previousSnapshot?.otfPricePerDevice,
            ),
          },
          subscriptionOnly: true,
          otfChanged: false,
          previousPlanMode: normalizeValue(currentAllocationRow?.planMode),
          previousPlanFee: normalizeValue(currentAllocationRow?.planFee),
          newPlanMode: effectivePlanLabel,
          newPlanFee: nextPlanFee,
          effectiveDate: effectiveDateValue,
          remarks,
          scope: "Allocation Model Group",
          validatedAt: allocationModelValidation.validatedAt,
          savedAndSubmittedAt: submittedAt,
          submissionStatus: "Saved & Submitted",
        };

        return {
          ...record,
          allocationRows: nextAllocationRows,
          allocationChainNetwork:
            getAllocationChainName(selectedModel) ||
            getAllocationChainName(record),
          allocationSettingsSnapshot: {
            ...selectedModel,
            otfPricePerDevice: normalizeValue(
              record?.allocationSettingsSnapshot?.otfPricePerDevice,
            ),
            chainName:
              getAllocationChainName(selectedModel) ||
              getAllocationChainName(record),
            allocationPurpose:
              getAllocationPurpose(selectedModel) ||
              normalizeValue(
                record?.allocationSettingsSnapshot?.allocationPurpose,
              ),
            modelId: selectedModel.modelId,
            modelName: selectedModel.name,
            effectiveFrom: effectiveDateValue,
          },
          allocationModelEffectiveDate: effectiveDateValue,
          allocationModelChangeStatus: "Effective",
          allocationModelLastSubmittedAt: submittedAt,
          allocationModelHistory: [
            ...(Array.isArray(record?.allocationModelHistory)
              ? record.allocationModelHistory
              : []),
            historyEntry,
          ],
        };
      }),
    );

    setAllocationModelValidation(null);
    setAllocationModelChangeDraft({
      modelId: "",
      effectiveDate: "",
      remarks: "",
      planSelections: {},
    });
    setPriceChangeMessage(
      "Allocation Model change saved and submitted successfully. The validated Allocation, Plan and Subscription Fee are now effective; OTF remains unchanged.",
    );
    setActiveRecordId("");
  }

  async function persistOpenPriceChangeWorkItem(
    record,
    sourceBillingRecordId,
    requestId,
    initiatedAt,
    currentFee,
  ) {
    const isCommon = isCommonComplexBillingRecord(record);
    const groupKey = isCommon ? getCommonComplexBillingGroupKey(record) : "";
    const affectedRecords = stage3Records.filter((candidate) =>
      isCommon
        ? getCommonComplexBillingGroupKey(candidate) === groupKey &&
          !isFoCOriginRecord(candidate)
        : candidate.recordId === record.recordId,
    );

    if (affectedRecords.some((candidate) => !normalizeValue(candidate.backendSiteId))) {
      throw new Error("A canonical backend Site ID is required to open Price Change & Reallocation.");
    }

    await Promise.all(
      affectedRecords.map((candidate) => {
        const currentStage3 =
          candidate?.stage3Data && typeof candidate.stage3Data === "object"
            ? candidate.stage3Data
            : candidate;
        const request = {
          requestId,
          status: "Pending Validation",
          approvalStatus: "Not Required",
          makerCheckerReady: true,
          sourceBillingRecordId,
          initiatedAt,
          effectiveDate: "",
          remarks: "",
          previousFee: normalizeValue(currentFee),
          newFee: normalizeValue(currentFee),
          changeType: "Price Change",
          previousMode: normalizeValue(candidate?.subscriptionMode),
          newMode: normalizeValue(candidate?.subscriptionMode),
          scope: isCommon ? "Common Complex" : "Site",
        };

        return billingApiRequest(
          `/sites/${encodeURIComponent(candidate.backendSiteId)}`,
          {
            method: "POST",
            body: JSON.stringify({
              stage3Data: {
                ...currentStage3,
                priceChangeRequest: request,
              },
            }),
          },
        );
      }),
    );
  }

  async function openPriceChangeDraftForRecord(record, sourceBillingRecordId = "") {
    if (!record) return;
    setAllocationModelValidation(null);
    setAllocationValidationPopup(null);
    priceChangeTransactionIdRef.current = crypto.randomUUID();

    if (!isPriceChangeLifecycleActive(record)) {
      alert(
        "Closed, paused, removed, or inactive billing records are not eligible for Price Change & Reallocation.",
      );
      return;
    }

    const existingRequest = getPriceChangeRequest(record);
    const openRequest = isPriceChangeRequestOpen(record) ? existingRequest : null;
    const currentFee = isCommonComplexBillingRecord(record)
      ? resolveCommonComplexSubscriptionFee(
          record,
          stage3Records,
          commonComplexAllocations,
        )
      : resolveSiteWiseSubscriptionFee(record) ||
        normalizeValue(record.subscriptionFee);

    const initialChangeType =
      normalizeValue(openRequest?.changeType) || "Price Change";
    const currentMode = normalizeSubscriptionModeLabel(record?.subscriptionMode);
    const requestedMode = normalizeValue(openRequest?.newMode);
    const initialNewMode =
      initialChangeType === "Price Change"
        ? currentMode
        : requestedMode && normalizeSubscriptionModeLabel(requestedMode) !== currentMode
          ? normalizeSubscriptionModeLabel(requestedMode)
          : getDefaultAlternativeSubscriptionMode(currentMode);

    setPriceChangeDraft({
      changeType: initialChangeType,
      effectiveDate: normalizeValue(openRequest?.effectiveDate),
      remarks: "",
      newFee: normalizeValue(openRequest?.newFee) || normalizeValue(currentFee),
      newMode: initialNewMode,
    });
    setPriceChangeEffectiveDateDisplay(
      formatPriceChangeDateDisplay(openRequest?.effectiveDate),
    );

    setAllocationModelChangeDraft({
      modelId: "",
      effectiveDate: "",
      remarks: "",
      planSelections: {},
    });
    setPriceChangeMessage("");
    setActiveWorkspace(workspaceModes.BILLING);
    setActiveBillingSection(billingSections.REALLOCATION);
    setActiveRecordId(record.recordId);

    if (isCommonComplexBillingRecord(record)) {
      const groupKey = getCommonComplexBillingGroupKey(record);
      const allocationGroup = commonComplexAllocations[groupKey] || {};
      const snapshot = {};
      Object.entries(allocationGroup).forEach(([siteKey, entry]) => {
        snapshot[siteKey] = normalizeValue(entry?.allocatedFee);
      });
      setReallocationOriginalAllocations(snapshot);
    } else {
      setReallocationOriginalAllocations({});
    }

    if (sourceBillingRecordId && !openRequest) {
      const initiatedAt = new Date().toISOString();
      const isCommon = isCommonComplexBillingRecord(record);
      const groupKey = isCommon ? getCommonComplexBillingGroupKey(record) : "";
      const requestId = `price-change-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;

      try {
        await persistOpenPriceChangeWorkItem(
          record,
          sourceBillingRecordId,
          requestId,
          initiatedAt,
          currentFee,
        );
      } catch (error) {
        setPriceChangeMessage(
          `Price Change work item could not be opened. ${error.message}`,
        );
        setActiveRecordId("");
        setActiveBillingSection("");
        return;
      }

      setStage3Records((currentRecords) =>
        currentRecords.map((candidate) => {
          const affected = isCommon
            ? getCommonComplexBillingGroupKey(candidate) === groupKey &&
              !isFoCOriginRecord(candidate)
            : candidate.recordId === record.recordId;
          if (!affected) return candidate;

          return {
            ...candidate,
            priceChangeRequest: {
              requestId,
              status: "Pending Validation",
              approvalStatus: "Not Required",
              makerCheckerReady: true,
              sourceBillingRecordId,
              initiatedAt,
              effectiveDate: "",
              remarks: "",
              previousFee: normalizeValue(currentFee),
              newFee: normalizeValue(currentFee),
              changeType: "Price Change",
              previousMode: normalizeValue(candidate?.subscriptionMode),
              newMode: normalizeValue(candidate?.subscriptionMode),
              scope: isCommon ? "Common Complex" : "Site",
            },
          };
        }),
      );
    }

    window.requestAnimationFrame(() =>
      guidedScrollToElement(priceChangeReallocationRef.current),
    );
  }

  function handleInitiatePriceChangeFromBillingRecord(billingRecord) {
    const stageRecord = getStage3RecordForPriceChangeBillingRecord(billingRecord);

    if (!stageRecord) {
      alert("Unable to locate the active Stage 3 site for this Billing Record.");
      return;
    }

    if (
      !isPriceChangeLifecycleActive(billingRecord) ||
      !isPriceChangeLifecycleActive(stageRecord)
    ) {
      alert(
        "This Billing Record is Closed / Paused / Inactive and cannot be moved to Price Change & Reallocation.",
      );
      return;
    }

    openPriceChangeDraftForRecord(
      stageRecord,
      normalizeValue(billingRecord.billingRecordId),
    );
  }

  function getPriceChangeNormalizedMode(record) {
    return normalizeValue(record?.subscriptionMode)
      .toLowerCase()
      .replace(/[-_]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function getProcessedBillingRecordsForPriceChange(record) {
    const screenCode = normalizeValue(record?.screenCode);
    const sourceRecordId = normalizeValue(record?.recordId);

    return billingRecords
      .filter((billingRecord) => {
        const submitted = ["Submitted to Billing Team", "Submitted", "Sent to Billing Team"].includes(
          normalizeValue(billingRecord?.submissionStatus),
        );
        if (!submitted) return false;

        return (
          (sourceRecordId && normalizeValue(billingRecord?.sourceRecordId) === sourceRecordId) ||
          (screenCode && normalizeValue(billingRecord?.screenCode || billingRecord?.siteScope) === screenCode)
        );
      })
      .sort((left, right) => {
        const leftDate = parseLocalDateValue(left?.billingPeriodFrom);
        const rightDate = parseLocalDateValue(right?.billingPeriodFrom);
        return (rightDate?.getTime() || 0) - (leftDate?.getTime() || 0);
      });
  }

  function getPriceChangeErpInvoiceEntries(record) {
    const recordId = normalizeValue(record?.recordId);
    const screenCode = normalizeValue(record?.screenCode).toUpperCase();
    const sources = [
      ...billingRecords.filter((billingRecord) =>
        (recordId && normalizeValue(billingRecord?.sourceRecordId) === recordId) ||
        (screenCode && normalizeValue(billingRecord?.screenCode || billingRecord?.siteScope).toUpperCase() === screenCode),
      ),
      record?.firstTimeBillingRecord,
      record,
    ].filter(Boolean);

    return sources.flatMap((source) => {
      const entries = Array.isArray(source?.invoiceEntries)
        ? source.invoiceEntries
        : [source];
      return entries.map((entry) => ({ entry, source }));
    });
  }

  function isValidPriceChangeErpInvoice(entry, source) {
    const invoiceNumber = normalizeValue(
      entry?.invoiceNumber || entry?.erpInvoiceNumber || source?.invoiceNumber || source?.erpInvoiceNumber,
    );
    if (!invoiceNumber) return false;

    const invoiceState = normalizeValue(
      entry?.erpInvoiceStatus ||
        entry?.erpInvoiceState ||
        source?.erpInvoiceStatus ||
        source?.erpInvoiceState ||
        entry?.invoiceStatus ||
        source?.invoiceStatus,
    ).toLowerCase();

    return ![
      "cancelled",
      "canceled",
      "revoked",
      "void",
      "voided",
      "reversed",
    ].some((state) => invoiceState.includes(state));
  }

  function getPriceChangeInvoiceLock(record, effectiveDate) {
    if (!record || !effectiveDate) return null;
    const isCommon = isCommonComplexBillingRecord(record);
    const groupKey = isCommon ? getCommonComplexBillingGroupKey(record) : "";
    const affectedRecords = isCommon
      ? stage3Records.filter(
          (candidate) => getCommonComplexBillingGroupKey(candidate) === groupKey,
        )
      : [record];

    for (const affectedRecord of affectedRecords) {
      const cycleBounds = getPriceChangeCycleBounds(affectedRecord, effectiveDate);
      if (!cycleBounds) continue;
      const normalizedMode = getPriceChangeNormalizedMode(affectedRecord);
      const monthly = ["", "monthly", "month"].includes(normalizedMode);
      const lock = getPriceChangeErpInvoiceEntries(affectedRecord).find(({ entry, source }) => {
        if (!isValidPriceChangeErpInvoice(entry, source)) return false;
        const from = parseLocalDateValue(
          entry?.periodFrom || source?.billingPeriodFrom || entry?.billingDate,
        );
        const to = parseLocalDateValue(
          entry?.periodTo || source?.billingPeriodTo || entry?.periodFrom || entry?.billingDate,
        );
        if (!from || !to || to < from) return false;

        return monthly
          ? from.getFullYear() === effectiveDate.getFullYear() &&
              from.getMonth() === effectiveDate.getMonth()
          : from <= cycleBounds.cycleEnd && to >= cycleBounds.cycleStart;
      });
      if (lock) return lock;
    }

    return null;
  }

  function getLatestProcessedPeriodEndForPriceChange(record) {
    const normalizedMode = getPriceChangeNormalizedMode(record);
    const isMonthly = ["", "monthly", "month"].includes(normalizedMode);
    const processedRecords = getProcessedBillingRecordsForPriceChange(record);

    return processedRecords
      .map((billingRecord) => {
        const periodFrom = parseLocalDateValue(billingRecord?.billingPeriodFrom);
        const periodTo = parseLocalDateValue(billingRecord?.billingPeriodTo);

        if (isMonthly && periodFrom) {
          // Monthly price changes are blocked only through the month that was
          // actually processed. A legacy First Time Billing snapshot can carry
          // a later periodTo; never let that incorrectly block the next month.
          return new Date(periodFrom.getFullYear(), periodFrom.getMonth() + 1, 0);
        }

        return periodTo;
      })
      .filter(Boolean)
      .sort((left, right) => right - left)[0] || null;
  }

  function getMinimumPriceChangeEffectiveDate(record) {
    if (!record) return "";

    const normalizedMode = getPriceChangeNormalizedMode(record);
    const isMonthly = ["", "monthly", "month"].includes(normalizedMode);
    const isCommon = isCommonComplexBillingRecord(record);
    const groupKey = isCommon ? getCommonComplexBillingGroupKey(record) : "";
    const affectedRecords = isCommon
      ? stage3Records.filter(
          (candidate) => getCommonComplexBillingGroupKey(candidate) === groupKey,
        )
      : [record];

    const billingStartDates = affectedRecords
      .map((candidate) => parseLocalDateValue(candidate?.billingStartDate))
      .filter(Boolean);
    let minimumDate = billingStartDates.sort((left, right) => left - right)[0] || null;

    // Monthly: once a month has actually been billed, a retrospective price
    // change for that processed month is not allowed. Start from the next day
    // after the last processed month.
    if (isMonthly) {
      const latestProcessedEnd = affectedRecords
        .map((candidate) => getLatestProcessedPeriodEndForPriceChange(candidate))
        .filter(Boolean)
        .sort((left, right) => right - left)[0] || null;

      if (latestProcessedEnd) {
        const dayAfterProcessedPeriod = new Date(latestProcessedEnd);
        dayAfterProcessedPeriod.setDate(dayAfterProcessedPeriod.getDate() + 1);
        if (!minimumDate || dayAfterProcessedPeriod > minimumDate) {
          minimumDate = dayAfterProcessedPeriod;
        }
      }
    }

    // Half-Yearly / Annual intentionally allow a change inside the currently
    // active, already-invoiced cycle so the system can calculate a credit or
    // additional amount for the remaining validity period.
    return minimumDate ? formatLocalDateValue(minimumDate) : "";
  }

  function getPriceChangeProcessedConflict(record, effectiveDate) {
    if (!record || !effectiveDate) return null;

    const normalizedMode = getPriceChangeNormalizedMode(record);
    const isMonthly = ["", "monthly", "month"].includes(normalizedMode);
    const processedRecords = getProcessedBillingRecordsForPriceChange(record);

    if (isMonthly) {
      const latestProcessedEnd = getLatestProcessedPeriodEndForPriceChange(record);
      return latestProcessedEnd && effectiveDate <= latestProcessedEnd
        ? { type: "Closed Monthly Period", billingRecord: processedRecords[0] || null }
        : null;
    }

    const today = new Date();
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const containingRecord = processedRecords.find((billingRecord) => {
      const from = parseLocalDateValue(billingRecord?.billingPeriodFrom);
      const to = parseLocalDateValue(billingRecord?.billingPeriodTo);
      return from && to && effectiveDate >= from && effectiveDate <= to;
    });

    // A Half-Yearly/Annual price change may be retrospective within the still
    // active validity cycle. A completely closed historical cycle remains locked.
    if (containingRecord) {
      const periodTo = parseLocalDateValue(containingRecord?.billingPeriodTo);
      if (periodTo && periodTo < todayStart) {
        return { type: "Closed Historical Cycle", billingRecord: containingRecord };
      }
      return null;
    }

    const historicalConflict = processedRecords.find((billingRecord) => {
      const to = parseLocalDateValue(billingRecord?.billingPeriodTo);
      return to && to < todayStart && effectiveDate <= to;
    });

    return historicalConflict
      ? { type: "Closed Historical Cycle", billingRecord: historicalConflict }
      : null;
  }

  function getPriceChangeCurrentFee(record) {
    if (!record) return null;

    return parseAmountValue(
      isCommonComplexBillingRecord(record)
        ? resolveCommonComplexSubscriptionFee(
            record,
            stage3Records,
            commonComplexAllocations,
          )
        : resolveSiteWiseSubscriptionFee(record) || record?.subscriptionFee,
    );
  }

  function getPriceChangeCycleBounds(record, effectiveDate) {
    if (!record || !effectiveDate) return null;

    const normalizedMode = getPriceChangeNormalizedMode(record);
    const billingStartDate = parseLocalDateValue(record?.billingStartDate);
    const effective = new Date(
      effectiveDate.getFullYear(),
      effectiveDate.getMonth(),
      effectiveDate.getDate(),
    );

    if (["", "monthly", "month"].includes(normalizedMode)) {
      let cycleStart = new Date(effective.getFullYear(), effective.getMonth(), 1);
      const cycleEnd = new Date(effective.getFullYear(), effective.getMonth() + 1, 0);

      // Preserve the true partial first month when Billing Start Date is mid-month.
      if (
        billingStartDate &&
        billingStartDate.getFullYear() === effective.getFullYear() &&
        billingStartDate.getMonth() === effective.getMonth() &&
        billingStartDate > cycleStart
      ) {
        cycleStart = new Date(billingStartDate);
      }

      return { cycleStart, cycleEnd };
    }

    if (!billingStartDate) return null;

    const isHalfYear = ["half year", "half yearly", "half-yearly", "semi annual"].includes(normalizedMode);
    const isAnnual = ["yearly", "annual", "one year"].includes(normalizedMode);
    if (!isHalfYear && !isAnnual) return null;

    let cycleStart = new Date(
      billingStartDate.getFullYear(),
      billingStartDate.getMonth(),
      billingStartDate.getDate(),
    );

    // Advance from the original Billing Start Date until we reach the cycle
    // containing the Effective Date. This keeps Half-Yearly / Annual cycles
    // anchored to the site's actual billing anniversary rather than FY dates.
    for (let guard = 0; guard < 240; guard += 1) {
      const nextCycleStart = isAnnual
        ? addCalendarYears(cycleStart, 1)
        : addCalendarMonths(cycleStart, 6);
      const cycleEnd = new Date(nextCycleStart);
      cycleEnd.setDate(cycleEnd.getDate() - 1);

      if (effective >= cycleStart && effective <= cycleEnd) {
        return { cycleStart, cycleEnd };
      }

      if (effective < cycleStart) return null;
      cycleStart = nextCycleStart;
    }

    return null;
  }

  function getAmountAlreadyBilledForPriceChangeCycle(record, effectiveDate, cycleStart, cycleEnd) {
    const normalizedMode = getPriceChangeNormalizedMode(record);
    const isMonthly = ["", "monthly", "month"].includes(normalizedMode);
    const processedRecords = getProcessedBillingRecordsForPriceChange(record);

    const matchedRecords = processedRecords.filter((billingRecord) => {
      const periodFrom = parseLocalDateValue(billingRecord?.billingPeriodFrom);
      const periodTo = parseLocalDateValue(billingRecord?.billingPeriodTo);
      if (!periodFrom) return false;

      if (isMonthly) {
        // A legacy First Time Billing snapshot may incorrectly carry a multi-month
        // periodTo. Monthly history belongs only to the month of periodFrom.
        return (
          periodFrom.getFullYear() === effectiveDate.getFullYear() &&
          periodFrom.getMonth() === effectiveDate.getMonth()
        );
      }

      return Boolean(periodTo && periodFrom <= cycleEnd && periodTo >= cycleStart);
    });

    return matchedRecords.reduce((total, billingRecord) => {
      const amount = parseAmountValue(
        billingRecord?.billingSubscriptionFee ??
          billingRecord?.billingAmountBeforeGST ??
          billingRecord?.invoiceAmountBeforeGST ??
          billingRecord?.actualSubscriptionFee,
      );
      return total + (amount ?? 0);
    }, 0);
  }

  function getPriceChangeAdjustmentPreview(record, draft = priceChangeDraft) {
    if (!record) return null;

    const effectiveDate = parseLocalDateValue(draft?.effectiveDate);
    const newFee = parseAmountValue(draft?.newFee);
    if (!effectiveDate || newFee === null) return null;

    const currentFee = getPriceChangeCurrentFee(record);
    if (currentFee === null) return null;

    const cycleBounds = getPriceChangeCycleBounds(record, effectiveDate);
    if (!cycleBounds) return null;

    const { cycleStart, cycleEnd } = cycleBounds;
    const amountAlreadyBilled = getAmountAlreadyBilledForPriceChangeCycle(
      record,
      effectiveDate,
      cycleStart,
      cycleEnd,
    );

    const totalCycleDays = getInclusiveDateDifference(cycleStart, cycleEnd) || 0;
    if (!totalCycleDays) return null;

    const oldEnd = new Date(effectiveDate);
    oldEnd.setDate(oldEnd.getDate() - 1);
    const oldDays = effectiveDate > cycleStart
      ? getInclusiveDateDifference(cycleStart, oldEnd) || 0
      : 0;

    const changeType = normalizeValue(draft?.changeType) || "Price Change";
    const modeChange = changeType === "Mode Change" || changeType === "Price + Mode Change";
    let oldEntitlement = (currentFee / totalCycleDays) * oldDays;
    let commonUnusedOldEntitlement = null;
    let commonUtilizedOldPeriod = null;

    if (isCommonComplexBillingRecord(record) && modeChange) {
      const commonRows = getCommonComplexBillingRows(
        record,
        stage3Records,
        commonComplexAllocations,
        commonScopeMembershipRecords,
      );

      let utilized = 0;
      commonRows.forEach((row) => {
        const rowRecord = row.relatedRecord;
        const oldAllocation = parseAmountValue(
          normalizeValue(reallocationOriginalAllocations?.[row.siteKey]) ||
            row.allocatedFee,
        );
        if (!rowRecord || oldAllocation === null) return;

        const rowBounds = getPriceChangeCycleBounds(rowRecord, effectiveDate);
        if (!rowBounds) return;

        const rowOldEnd = new Date(effectiveDate);
        rowOldEnd.setDate(rowOldEnd.getDate() - 1);
        const rowEffectiveEnd =
          rowOldEnd < rowBounds.cycleEnd ? rowOldEnd : rowBounds.cycleEnd;
        const rowDays =
          rowEffectiveEnd >= rowBounds.cycleStart
            ? getInclusiveDateDifference(rowBounds.cycleStart, rowEffectiveEnd) || 0
            : 0;
        const rowCycleDays =
          getInclusiveDateDifference(rowBounds.cycleStart, rowBounds.cycleEnd) || 0;

        if (rowCycleDays > 0) {
          utilized += (oldAllocation / rowCycleDays) * rowDays;
        }
      });

      commonUtilizedOldPeriod = utilized;
      commonUnusedOldEntitlement = Math.max(0, currentFee - utilized);
      oldEntitlement = utilized;
    }

    if (modeChange) {
      const newMode = normalizeValue(draft?.newMode);
      const normalizedNewMode = newMode
        .toLowerCase()
        .replace(/[-_]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
      const newCycleStart = new Date(effectiveDate);
      let newCycleEnd = null;
      let newCycleCharge = newFee;

      if (["", "monthly", "month"].includes(normalizedNewMode)) {
        newCycleEnd = new Date(
          newCycleStart.getFullYear(),
          newCycleStart.getMonth() + 1,
          0,
        );
        const monthDays = newCycleEnd.getDate();
        const billableDays = getInclusiveDateDifference(newCycleStart, newCycleEnd) || 0;
        newCycleCharge = monthDays
          ? (newFee / monthDays) * billableDays
          : newFee;
      } else if (["half year", "half yearly", "semi annual"].includes(normalizedNewMode)) {
        newCycleEnd = addCalendarMonths(newCycleStart, 6);
        newCycleEnd.setDate(newCycleEnd.getDate() - 1);
      } else if (["yearly", "annual", "one year"].includes(normalizedNewMode)) {
        newCycleEnd = addCalendarYears(newCycleStart, 1);
        newCycleEnd.setDate(newCycleEnd.getDate() - 1);
      }

      const newCycleFrom = formatLocalDateValue(newCycleStart);
      const newCycleTo = newCycleEnd ? formatLocalDateValue(newCycleEnd) : "";
      const oldAdjustment =
        commonUnusedOldEntitlement !== null
          ? newCycleCharge - commonUnusedOldEntitlement
          : oldEntitlement - amountAlreadyBilled;
      const netNextBilling =
        commonUnusedOldEntitlement !== null
          ? Math.max(0, oldAdjustment)
          : newCycleCharge + oldAdjustment;

      return {
        modeChange: true,
        commonGroupAdjustment: commonUnusedOldEntitlement !== null,
        previousCommonFee:
          commonUnusedOldEntitlement !== null
            ? formatAmountValue(currentFee)
            : "",
        utilizedOldPeriod:
          commonUtilizedOldPeriod !== null
            ? formatAmountValue(commonUtilizedOldPeriod)
            : "",
        unusedOldEntitlement:
          commonUnusedOldEntitlement !== null
            ? formatAmountValue(commonUnusedOldEntitlement)
            : "",
        cycleFrom: formatLocalDateValue(cycleStart),
        cycleTo: formatLocalDateValue(cycleEnd),
        oldModeEnd: formatLocalDateValue(oldEnd),
        amountAlreadyBilled: formatAmountValue(amountAlreadyBilled),
        revisedEntitlement: formatAmountValue(oldEntitlement),
        adjustmentAmount: formatAmountValue(Math.abs(oldAdjustment)),
        adjustmentSigned: oldAdjustment,
        adjustmentType: oldAdjustment < -0.005
          ? "Credit / Excess"
          : oldAdjustment > 0.005
            ? "Balance to Collect"
            : "No Adjustment",
        newCycleFrom,
        newCycleTo,
        newCycleFee: formatAmountValue(newCycleCharge),
        netNextBilling: formatAmountValue(Math.max(0, netNextBilling)),
      };
    }

    const newDays = effectiveDate <= cycleEnd
      ? getInclusiveDateDifference(effectiveDate, cycleEnd) || 0
      : 0;
    const revisedEntitlement =
      oldEntitlement + (newFee / totalCycleDays) * newDays;
    const adjustment = revisedEntitlement - amountAlreadyBilled;

    return {
      modeChange: false,
      cycleFrom: formatLocalDateValue(cycleStart),
      cycleTo: formatLocalDateValue(cycleEnd),
      oldModeEnd: formatLocalDateValue(oldEnd),
      amountAlreadyBilled: formatAmountValue(amountAlreadyBilled),
      revisedEntitlement: formatAmountValue(revisedEntitlement),
      adjustmentAmount: formatAmountValue(Math.abs(adjustment)),
      adjustmentSigned: adjustment,
      adjustmentType: adjustment < -0.005
        ? "Credit / Excess"
        : adjustment > 0.005
          ? "Balance to Collect"
          : "No Adjustment",
      newCycleFrom: "",
      newCycleTo: "",
      newCycleFee: "",
      netNextBilling: "",
    };
  }

  function handleCancelPriceChange() {
    setAllocationModelValidation(null);
    setAllocationValidationPopup(null);
    priceChangeTransactionIdRef.current = "";
    if (!activeRecord) {
      setPriceChangeDraft({ changeType: "Price Change", effectiveDate: "", remarks: "", newFee: "", newMode: "" });
      setPriceChangeEffectiveDateDisplay("");
      setPriceChangeMessage("");
      return;
    }

    const isCommon = isCommonComplexBillingRecord(activeRecord);
    const groupKey = isCommon ? getCommonComplexBillingGroupKey(activeRecord) : "";

    if (isCommon && groupKey) {
      setCommonComplexAllocations((currentAllocations) => {
        const currentGroup = currentAllocations[groupKey] || {};
        const nextGroup = { ...currentGroup };

        Object.entries(reallocationOriginalAllocations || {}).forEach(
          ([siteKey, allocatedFee]) => {
            nextGroup[siteKey] = {
              ...(nextGroup[siteKey] || {}),
              allocatedFee: normalizeValue(allocatedFee),
            };
          },
        );

        return {
          ...currentAllocations,
          [groupKey]: nextGroup,
        };
      });
    }

    setStage3Records((currentRecords) =>
      currentRecords.map((record) => {
        const affected = isCommon
          ? getCommonComplexBillingGroupKey(record) === groupKey
          : record.recordId === activeRecord.recordId;

        if (!affected) return record;

        const request = getPriceChangeRequest(record);
        if (normalizeValue(request?.status) !== "Pending Validation") {
          return record;
        }

        const nextRecord = { ...record };
        delete nextRecord.priceChangeRequest;
        return nextRecord;
      }),
    );

    setPriceChangeDraft({ changeType: "Price Change", effectiveDate: "", remarks: "", newFee: "", newMode: "" });
    setPriceChangeEffectiveDateDisplay("");
    setPriceChangeMessage("");
    setReallocationOriginalAllocations({});
    setAllocationModelChangeDraft({ modelId: "", effectiveDate: "", remarks: "" });
    setActiveRecordId("");
  }

  function handleGoToPriceChangeAllocation() {
    setAllocationValidationPopup(null);
    window.requestAnimationFrame(() => {
      guidedScrollToElement(priceChangeAllocationRef.current);
      priceChangeAllocationRef.current
        ?.querySelector("input")
        ?.focus({ preventScroll: true });
    });
  }

  async function handleValidatePriceChange() {
    if (!activeRecord || priceChangeApplyPendingRef.current) return;

    const transactionId =
      priceChangeTransactionIdRef.current || crypto.randomUUID();
    priceChangeTransactionIdRef.current = transactionId;
    const changeType = normalizeValue(priceChangeDraft.changeType) || "Price Change";
    const effectiveDateValue = normalizeValue(priceChangeDraft.effectiveDate);
    const remarks = normalizeValue(priceChangeDraft.remarks);
    const newFeeValue = sanitizeAmount(priceChangeDraft.newFee);
    const newFee = parseAmountValue(newFeeValue);
    const previousMode = normalizeValue(activeRecord.subscriptionMode) || "Monthly";
    const newMode = normalizeValue(priceChangeDraft.newMode) || previousMode;
    const modeChange = changeType === "Mode Change" || changeType === "Price + Mode Change";
    const normalizedPreviousMode = previousMode
      .toLowerCase()
      .replace(/[-_]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    const currentModeIsMonthly = ["", "monthly", "month"].includes(normalizedPreviousMode);
    const currentFee = getPriceChangeCurrentFee(activeRecord);

    if (currentModeIsMonthly && changeType === "Price + Mode Change") {
      setPriceChangeMessage(
        "Price + Mode Change is not applicable when the current Subscription Mode is Monthly. Use Mode Change and enter the New Subscription Fee.",
      );
      return;
    }

    if (
      !effectiveDateValue ||
      !isValidFourDigitYear(effectiveDateValue) ||
      !isValidDateValue(effectiveDateValue)
    ) {
      setPriceChangeMessage(
        "Please enter a valid Effective Date using a 4-digit year.",
      );
      return;
    }

    if (!remarks) {
      setPriceChangeMessage("Remarks / Reason is required.");
      return;
    }

    if (newFee === null || newFee < 0) {
      setPriceChangeMessage("Enter a valid New Subscription Fee.");
      return;
    }

    if (changeType === "Price Change" && currentFee !== null && Math.abs(newFee - currentFee) < 0.005) {
      setPriceChangeMessage("New Subscription Fee must be different from the Current Subscription Fee for Price Change.");
      return;
    }

    if (changeType === "Price + Mode Change" && currentFee !== null && Math.abs(newFee - currentFee) < 0.005) {
      setPriceChangeMessage("Price + Mode Change requires both a different New Subscription Mode and a different New Subscription Fee.");
      return;
    }

    if (
      modeChange &&
      (!newMode ||
        normalizeSubscriptionModeLabel(newMode) ===
          normalizeSubscriptionModeLabel(previousMode))
    ) {
      setPriceChangeMessage(
        "New Subscription Mode cannot be the same as Current Subscription Mode.",
      );
      return;
    }

    const isCommon = isCommonComplexBillingRecord(activeRecord);
    const groupKey = isCommon ? getCommonComplexBillingGroupKey(activeRecord) : "";
    const commonRows = isCommon
      ? getCommonComplexBillingRows(
          activeRecord,
          stage3Records,
          commonComplexAllocations,
          commonScopeMembershipRecords,
        )
      : [];

    if (isCommon) {
      const allocationWasUpdated = commonRows.some(
        (row) =>
          normalizeValue(row.allocatedFee) !==
          normalizeValue(reallocationOriginalAllocations[row.siteKey]),
      );
      if (!allocationWasUpdated) {
        setAllocationValidationPopup({
          title: "Allocation Required",
          message:
            "Please update the site-wise allocation before applying the price change. The total allocated amount must equal the New Common Subscription Fee.",
        });
        return;
      }

      const incompleteRows = commonRows.filter(
        (row) => parseAmountValue(row.allocatedFee) === null,
      );
      if (incompleteRows.length > 0) {
        setAllocationValidationPopup({
          title: "Allocation Incomplete",
          message:
            "Please enter an allocation amount for all applicable sites before applying the price change.",
        });
        return;
      }

      const allocationTotal = commonRows.reduce(
        (total, row) => total + parseAmountValue(row.allocatedFee),
        0,
      );
      const difference = newFee - allocationTotal;
      if (Math.round(allocationTotal * 100) !== Math.round(newFee * 100)) {
        setAllocationValidationPopup({
          title: "Allocation Total Mismatch",
          message:
            "The site-wise allocation total does not match the New Common Subscription Fee.",
          newFee: formatAmountValue(newFee),
          allocationTotal: formatAmountValue(allocationTotal),
          difference: formatAmountValue(Math.abs(difference)),
          differenceMessage:
            difference > 0
              ? `₹${formatAmountValue(difference)} remains to be allocated.`
              : `Allocation exceeds the New Common Subscription Fee by ₹${formatAmountValue(Math.abs(difference))}.`,
        });
        return;
      }
    }

    const effectiveDate = parseLocalDateValue(effectiveDateValue);
    const billingStartDate = parseLocalDateValue(activeRecord.billingStartDate);

    if (billingStartDate && effectiveDate < billingStartDate) {
      setPriceChangeMessage(
        "Price Change Effective Date cannot be earlier than the Billing Start Date.",
      );
      return;
    }

    const invoiceLock = getPriceChangeInvoiceLock(activeRecord, effectiveDate);
    if (invoiceLock && changeType !== "Mode Change") {
      setPriceChangeMessage(
        "Price and Effective Date are locked because the affected billing period has a valid ERP invoice. The change can be made only after ERP sync confirms the invoice is revoked or cancelled and no replacement invoice exists.",
      );
      return;
    }

    const affectedRecords = isCommon
      ? stage3Records.filter(
          (record) => getCommonComplexBillingGroupKey(record) === groupKey,
        )
      : [activeRecord];

    const processedConflict = affectedRecords
      .map((record) => ({
        record,
        conflict: getPriceChangeProcessedConflict(record, effectiveDate),
      }))
      .find((entry) => entry.conflict);

    if (processedConflict) {
      setPriceChangeMessage(
        processedConflict.conflict?.type === "Closed Monthly Period"
          ? "The Effective Date falls within a Monthly billing period that has already been processed. Select the next unbilled date/month; historical Billing Records will not be overwritten."
          : "The Effective Date falls within a fully closed historical billing cycle. Select a date in the current/open validity period; historical Billing Records will not be overwritten.",
      );
      return;
    }

    const validatedAt = new Date().toISOString();
    const previousCommonFee = isCommon
      ? parseAmountValue(
          resolveCommonComplexSubscriptionFee(
            activeRecord,
            stage3Records,
            commonComplexAllocations,
          ),
        )
      : null;

    const allocationGroup = isCommon
      ? commonComplexAllocations[groupKey] || {}
      : {};
    let appliedPriceChange = null;

    if (isCommon) {
      const items = commonRows.map((row) => {
        const matchingRecord =
          row.relatedRecord ||
          stage3Records.find(
            (record) =>
              getCommonComplexBillingSiteKey(record.screenCode) === row.siteKey,
          ) ||
          stage2Records.find(
            (record) =>
              getCommonComplexBillingSiteKey(record.screenCode) === row.siteKey,
          );
        return {
          siteId: normalizeValue(matchingRecord?.backendSiteId),
          screenCode: row.screenCode,
          allocatedFee: parseAmountValue(row.allocatedFee),
        };
      });
      const missingIdentity = items.find((item) => !item.siteId || !item.screenCode);
      if (missingIdentity) {
        setPriceChangeMessage(
          "A canonical backend Site ID is missing from the selected Common scope. No price change was applied.",
        );
        return;
      }

      priceChangeApplyPendingRef.current = true;
      setIsApplyingPriceChange(true);
      try {
        appliedPriceChange = await billingApiRequest("/stage3/allocation-reallocation/common-price-change/apply", {
          method: "POST",
          body: JSON.stringify({
            transactionId,
            billingCode: normalizeValue(activeRecord.billingCode),
            complexCode: normalizeValue(activeRecord.complexCode),
            changeType,
            newFee,
            newMode: modeChange ? newMode : previousMode,
            effectiveDate: effectiveDateValue,
            remarks,
            changedBy: "Operations",
            items,
          }),
        });
      } catch (error) {
        setPriceChangeMessage(
          `Price Change could not be applied. No values became active. ${error.message}`,
        );
        return;
      } finally {
        priceChangeApplyPendingRef.current = false;
        setIsApplyingPriceChange(false);
      }
    }

    const buildAppliedPriceChangeRecord = (record) => {
      const affected = isCommon
        ? getCommonComplexBillingGroupKey(record) === groupKey
        : record.recordId === activeRecord.recordId;
      if (!affected) return record;

      const siteKey = getCommonComplexBillingSiteKey(record.screenCode);
      const previousSiteFee = isCommon
        ? normalizeValue(reallocationOriginalAllocations[siteKey]) ||
          normalizeValue(allocationGroup?.[siteKey]?.allocatedFee)
        : resolveSiteWiseSubscriptionFee(record) || normalizeValue(record.subscriptionFee);
      const newSiteFee = isCommon
        ? normalizeValue(allocationGroup?.[siteKey]?.allocatedFee)
        : formatAmountValue(newFee);
      const previousFee = isCommon
        ? formatAmountValue(previousCommonFee)
        : previousSiteFee;

      const historyEntry = {
        transactionId,
        historyId: transactionId,
        previousFee,
        newFee: formatAmountValue(newFee),
        previousSiteFee,
        newSiteFee,
        effectiveDate: effectiveDateValue,
        remarks,
        changeType,
        previousMode: normalizeValue(record.subscriptionMode),
        newMode: modeChange ? newMode : normalizeValue(record.subscriptionMode),
        oldModeEndDate: modeChange
          ? (() => {
              const date = new Date(effectiveDate);
              date.setDate(date.getDate() - 1);
              return formatLocalDateValue(date);
            })()
          : "",
        scope: isCommon ? "Common Complex" : "Site",
        validatedAt,
        approvalStatus: "Not Required",
        makerCheckerReady: true,
      };

      const canonicalSite = appliedPriceChange?.sites?.find(
        (site) => normalizeValue(site?.id) === normalizeValue(record.backendSiteId),
      );
      const nextRecord = {
        ...record,
        ...(canonicalSite?.stage3Data || {}),
        subscriptionFee: formatAmountValue(newFee),
        subscriptionMode:
          modeChange && effectiveDate <= new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate())
            ? newMode
            : record.subscriptionMode,
        modeChangeEffectiveDate: modeChange ? effectiveDateValue : normalizeValue(record.modeChangeEffectiveDate),
        modeChangeTargetMode: modeChange ? newMode : normalizeValue(record.modeChangeTargetMode),
        priceChangeOverrideFee: isCommon ? normalizeValue(record.priceChangeOverrideFee) : formatAmountValue(newFee),
        priceChangeRequest: {
          ...(getPriceChangeRequest(record) || {}),
          status: "Effective",
          approvalStatus: "Not Required",
          makerCheckerReady: true,
          transactionId,
          effectiveDate: effectiveDateValue,
          remarks,
          changeType,
          previousMode: normalizeValue(record.subscriptionMode),
          newMode: modeChange ? newMode : normalizeValue(record.subscriptionMode),
          previousFee,
          newFee: formatAmountValue(newFee),
          validatedAt,
        },
        priceChangeHistory: [
          ...(Array.isArray(record.priceChangeHistory)
            ? record.priceChangeHistory
            : []),
          historyEntry,
        ],
      };
      return isCommon
        ? invalidateFirstTimeBillingValidation(nextRecord, "subscriptionFee")
        : nextRecord;
    };

    const nextStage3Records = stage3Records.map(buildAppliedPriceChangeRecord);

    if (!isCommon) {
      const nextRecord = nextStage3Records.find(
        (record) => record.recordId === activeRecord.recordId,
      );
      try {
        await billingApiRequest(
          `/stage3/allocation-reallocation/${encodeURIComponent(activeRecord.backendSiteId)}/price-change/apply`,
          {
            method: "POST",
            body: JSON.stringify({
              transactionId,
              screenCode: normalizeValue(activeRecord.screenCode),
              changeType,
              newFee,
              newMode: modeChange ? newMode : normalizeValue(activeRecord.subscriptionMode),
              effectiveDate: effectiveDateValue,
              remarks,
              changedBy: "Operations",
            }),
          },
        );
      } catch (error) {
        setPriceChangeMessage(
          `Price Change could not be saved. No values became active. ${error.message}`,
        );
        return;
      }
    }

    setStage3Records(nextStage3Records);

    if (isCommon) {
      setCommonComplexAllocations((currentAllocations) => {
        const currentGroup = currentAllocations[groupKey] || {};
        const nextGroup = {};
        Object.entries(currentGroup).forEach(([siteKey, entry]) => {
          nextGroup[siteKey] = {
            ...entry,
            locked: true,
            billingAllocationUpdatedAt: getCurrentLocalTimestamp(),
          };
        });
        return {
          ...currentAllocations,
          [groupKey]: nextGroup,
        };
      });
      const affectedSiteIds = new Set(
        (appliedPriceChange?.sites || commonRows.map((row) => row.relatedRecord))
          .map((site) => normalizeValue(site?.id || site?.backendSiteId))
          .filter(Boolean),
      );
      setStage2Records((currentRecords) =>
        currentRecords.map((record) =>
          affectedSiteIds.has(normalizeValue(record.backendSiteId))
            ? {
                ...record,
                subscriptionFee: formatAmountValue(newFee),
                subscriptionMode: modeChange ? newMode : record.subscriptionMode,
                firstTimeBillingValidationStatus: "Billing Calculation",
                firstTimeBillingValidatedAt: "",
                firstTimeBillingValidatedBy: "",
              }
            : record,
        ),
      );
    }

    const successMessage =
      "Price Change applied successfully. The previous commercial terms end one day before the Effective Date, and the new commercial terms take effect on the Effective Date. Any credit/balance from the previous billing cycle will be retained for billing adjustment.";
    setPriceChangeMessage(successMessage);
    window.alert(successMessage);
    priceChangeTransactionIdRef.current = "";
    setPriceChangeDraft({
      changeType: "Price Change",
      effectiveDate: "",
      remarks: "",
      newFee: "",
      newMode: "",
    });
    setPriceChangeEffectiveDateDisplay("");
    setReallocationOriginalAllocations({});
    setAllocationModelChangeDraft({
      modelId: "",
      effectiveDate: "",
      remarks: "",
      planSelections: {},
    });
    setActiveRecordId("");
    setActiveBillingSection("");
  }

  const derivedPriceChangeRows = useMemo(() => {
    const rows = [];
    const seenCommonGroups = new Set();

    stage3Records.forEach((record) => {
      if (!record || isFoCOriginRecord(record)) return;
      if (!isPriceChangeLifecycleActive(record)) return;
      if (!isPriceChangeRequestOpen(record)) return;

      // Finalized workflow rule:
      // Only an explicitly initiated open work item belongs here. General
      // post-billing eligibility does not create an operational PCR row.
      if (!hasFirstBillingStarted(record)) return;

      const currentFee = parseAmountValue(record.subscriptionFee);
      if (currentFee === null) return;

      if (isCommonComplexBillingRecord(record)) {
        const groupKey = getCommonComplexBillingGroupKey(record);
        if (!groupKey || seenCommonGroups.has(groupKey)) return;
        seenCommonGroups.add(groupKey);

        const allocationGroup = commonComplexAllocations[groupKey] || {};
        const previousAllocationTotal = Object.values(allocationGroup).reduce(
          (sum, entry) => {
            const fee = parseAmountValue(entry?.allocatedFee);
            return sum + (fee === null ? 0 : fee);
          },
          0,
        );

        const openRequest = stage3Records.find(
          (candidate) =>
            getCommonComplexBillingGroupKey(candidate) === groupKey &&
            isPriceChangeRequestOpen(candidate),
        );

        const billableRows = getCommonComplexBillingRows(
          record,
          stage3Records,
          commonComplexAllocations,
          commonScopeMembershipRecords,
        );

        rows.push({
          rowKey: `COMMON::${groupKey}`,
          recordId: record.recordId,
          billingCode: normalizeValue(record.billingCode),
          complexCode: normalizeValue(record.complexCode),
          screenCode: "-",
          screenName: normalizeValue(record.billingName) || normalizeValue(record.screenName),
          changeType: "Complex Common Reallocation",
          previousFee: formatAmountValue(previousAllocationTotal),
          currentFee: formatAmountValue(currentFee),
          billingStartDate: normalizeValue(record.billingStartDate),
          affectedSites: billableRows.length,
          status: openRequest
            ? normalizeValue(openRequest?.priceChangeRequest?.status) || "Pending Validation"
            : "Ready for New Change",
          isCommon: true,
        });
        return;
      }

      const screenCode = normalizeValue(record.screenCode).toUpperCase();
      if (!screenCode) return;

      const previousBillingRecord = billingRecords
        .filter(
          (billingRecord) =>
            normalizeValue(billingRecord.screenCode).toUpperCase() === screenCode,
        )
        .sort((left, right) => {
          const leftTime = Date.parse(left.submittedAt || left.createdAt || left.lastUpdatedAt || "") || 0;
          const rightTime = Date.parse(right.submittedAt || right.createdAt || right.lastUpdatedAt || "") || 0;
          return rightTime - leftTime;
        })[0];

      if (!previousBillingRecord) return;

      const previousFee = parseAmountValue(
        previousBillingRecord?.billingCommercialLockSnapshot?.subscriptionFee ??
          previousBillingRecord?.billingSubscriptionFee ??
          previousBillingRecord?.subscriptionFee,
      );

      const openRequest = isPriceChangeRequestOpen(record);

      rows.push({
        rowKey: `SITE::${screenCode}`,
        recordId: record.recordId,
        billingCode: normalizeValue(record.billingCode),
        complexCode: normalizeValue(record.complexCode),
        screenCode: normalizeValue(record.screenCode),
        screenName: normalizeValue(record.screenName),
        changeType: normalizeValue(record.complexCode)
          ? "Site / Screen-wise Price Change"
          : "Standalone Price Change",
        previousFee: formatAmountValue(previousFee),
        currentFee: formatAmountValue(currentFee),
        billingStartDate: normalizeValue(record.billingStartDate),
        affectedSites: 1,
        status: openRequest
          ? normalizeValue(record?.priceChangeRequest?.status) || "Pending Validation"
          : "Ready for New Change",
        isCommon: false,
      });
    });

    const filtered = rows.filter((row) => {
      if (!normalizedPriceChangeSearch) return true;
      return [
        row.billingCode,
        row.complexCode,
        row.screenCode,
        row.screenName,
        row.changeType,
        row.status,
      ]
        .join(" ")
        .toLowerCase()
        .includes(normalizedPriceChangeSearch);
    });

    return filtered.sort((left, right) => {
      const valueFor = (row) => {
        if (!priceChangeSortKey) return row.status;
        return row?.[priceChangeSortKey];
      };
      const comparison = compareColumnValues(valueFor(left), valueFor(right));
      return priceChangeSortOrder === "asc" ? comparison : -comparison;
    });
  }, [
    billingRecords,
    commonComplexAllocations,
    commonScopeMembershipRecords,
    normalizedPriceChangeSearch,
    priceChangeSortKey,
    priceChangeSortOrder,
    stage3Records,
  ]);

  const priceChangeRows = useMemo(() => {
    return derivedPriceChangeRows;
  }, [derivedPriceChangeRows]);

  function handlePriceChangeColumnSort(key) {
    if (priceChangeSortKey === key) {
      setPriceChangeSortOrder((current) => (current === "asc" ? "desc" : "asc"));
      return;
    }
    setPriceChangeSortKey(key);
    setPriceChangeSortOrder("asc");
  }

  function handleOpenPriceChange(row) {
    const record = stage3Records.find((candidate) => candidate.recordId === row.recordId);
    if (!record) {
      return;
    }

    if (row.isCommon) {
      const groupKey = getCommonComplexBillingGroupKey(record);
      const allocationGroup = commonComplexAllocations[groupKey] || {};
      const snapshot = {};
      Object.entries(allocationGroup).forEach(([siteKey, entry]) => {
        snapshot[siteKey] = normalizeValue(entry?.allocatedFee);
      });
      setReallocationOriginalAllocations(snapshot);
    } else {
      setReallocationOriginalAllocations({});
    }

    setAllocationSaveMessage("");
    openPriceChangeDraftForRecord(record);
  }

  const commonComplexBillingRows = useMemo(
    () =>
      activeRecord
        ? getCommonComplexBillingRows(
            activeRecord,
            stage3Records,
            commonComplexAllocations,
            commonScopeMembershipRecords,
          )
        : [],
    [activeRecord, stage1Records, stage2Records, stage3Records, commonComplexAllocations],
  );
  const commonComplexBillingValidationError = useMemo(() => {
    return getCommonComplexBillingValidationError(
      activeRecord,
      stage3Records,
      commonComplexAllocations,
      commonScopeMembershipRecords,
    );
  }, [activeRecord, stage1Records, stage2Records, stage3Records, commonComplexAllocations]);
  const commonComplexBillingMode = useMemo(
    () =>
    resolveCommonComplexBillingStatus(
      activeRecord,
      stage3Records,
      commonComplexAllocations,
      commonScopeMembershipRecords,
    ),
    [activeRecord, stage1Records, stage2Records, stage3Records, commonComplexAllocations],
  );
  const commonComplexBillingSummary = useMemo(() => {
    if (!activeRecord || !isCommonComplexBillingRecord(activeRecord)) {
      return null;
    }

    const resolvedCommonSubscriptionFee = resolveCommonComplexSubscriptionFee(
      activeRecord,
      stage3Records,
      commonComplexAllocations,
    );
    const commonSubscriptionFee = parseAmountValue(
      resolvedCommonSubscriptionFee,
    );
    // Total Complex Sites is an identity/count metric and must include every
    // screen in the complex, including FoC screens. The allocation rows below
    // intentionally exclude FoC screens, so do not derive the total from
    // commonComplexBillingRows.
    const groupKey = getCommonComplexBillingGroupKey(activeRecord);
    const relatedStage3Records = stage3Records.filter(
      (candidate) => getCommonComplexBillingGroupKey(candidate) === groupKey,
    );
    const relatedStage2Records = stage2Records.filter(
      (candidate) => getCommonComplexBillingGroupKey(candidate) === groupKey,
    );
    const relatedStage1Records = stage1Records.filter(
      (candidate) => getCommonComplexBillingGroupKey(candidate) === groupKey,
    );
    const allComplexSiteKeys = new Set();
    [
      ...(Array.isArray(activeRecord?.stage1GroupRows)
        ? activeRecord.stage1GroupRows
        : []),
      ...relatedStage3Records.flatMap((candidate) =>
        Array.isArray(candidate?.stage1GroupRows)
          ? candidate.stage1GroupRows
          : [],
      ),
      ...relatedStage3Records,
      ...relatedStage2Records,
      ...relatedStage1Records,
    ].forEach((row) => {
      const siteKey = getCommonComplexBillingSiteKey(row?.screenCode);
      if (siteKey) allComplexSiteKeys.add(siteKey);
    });

    const totalComplexSites = allComplexSiteKeys.size;
    const currentlyBillableSites = commonComplexBillingRows.filter(
      (row) => row.isBillable,
    ).length;
    // Pending Sites belongs to the commercially applicable allocation set,
    // not to the full complex count. FoC screens are excluded by design.
    const pendingSites = commonComplexBillingRows.length - currentlyBillableSites;
    let allocationTotal = 0;

    if (commonSubscriptionFee === null) {
      return {
        billingCode: normalizeValue(activeRecord.billingCode),
        complexCode: normalizeValue(activeRecord.complexCode),
        billingMode: commonComplexBillingMode,
        commonSubscriptionFee: normalizeValue(resolvedCommonSubscriptionFee),
        totalComplexSites,
        currentlyBillableSites,
        pendingSites,
        allocationTotal: "",
        difference: "",
      };
    }

    for (const row of commonComplexBillingRows) {
      const allocatedFee = parseAmountValue(row.allocatedFee);

      if (allocatedFee === null) {
        return {
          billingCode: normalizeValue(activeRecord.billingCode),
          complexCode: normalizeValue(activeRecord.complexCode),
          billingMode: commonComplexBillingMode,
          commonSubscriptionFee: normalizeValue(resolvedCommonSubscriptionFee),
          totalComplexSites,
          currentlyBillableSites,
          pendingSites,
          allocationTotal: "",
          difference: "",
        };
      }

      allocationTotal += allocatedFee;
    }

    const difference = commonSubscriptionFee - allocationTotal;

    return {
      billingCode: normalizeValue(activeRecord.billingCode),
      complexCode: normalizeValue(activeRecord.complexCode),
      billingMode: commonComplexBillingMode,
      commonSubscriptionFee: formatAmountValue(commonSubscriptionFee),
      totalComplexSites,
      currentlyBillableSites,
      pendingSites,
      allocationTotal: formatAmountValue(allocationTotal),
      difference: formatAmountValue(difference),
    };
  }, [
    activeRecord,
    stage1Records,
    stage3Records,
    stage2Records,
    commonComplexAllocations,
    commonComplexBillingRows,
    commonComplexBillingMode,
  ]);

  const activePriceChangeAdjustmentPreview = useMemo(
    () =>
      activeRecord && activeBillingSection === billingSections.REALLOCATION
        ? getPriceChangeAdjustmentPreview(activeRecord, priceChangeDraft)
        : null,
    [
      activeRecord,
      activeBillingSection,
      priceChangeDraft,
      billingRecords,
      stage3Records,
      commonComplexAllocations,
    ],
  );

  const activeCommonComplexBillingGroupKey = useMemo(
    () => (activeRecord ? getCommonComplexBillingGroupKey(activeRecord) : ""),
    [activeRecord],
  );
  const activeCommonComplexBillingRow = useMemo(() => {
    if (!activeRecord || !activeCommonComplexBillingGroupKey) {
      return null;
    }

    return commonComplexBillingRows.find(
      (row) => row.siteKey === activeAllocationSiteKey,
    ) || null;
  }, [
    activeAllocationSiteKey,
    activeCommonComplexBillingGroupKey,
    activeRecord,
    commonComplexBillingRows,
  ]);
  const submittedBillingRecords = useMemo(() => {
    const sixHoursMs = 6 * 60 * 60 * 1000;
    const nowMs = Date.now();

    const sourceRecords = canonicalBillingRecordRows || billingRecords;

    return dedupeBillingRecords(
      sourceRecords.filter(
        (record) =>
          !record?.isCombinedCommonBilling &&
          [
            "Submitted to Billing Team",
            "Sent to Billing Team",
          ].includes(normalizeValue(record.submissionStatus)),
      ),
      )
      .map((record) => {
        const sourceRecordId = normalizeValue(record.sourceRecordId);
        const screenCode = normalizeValue(record.screenCode || record.siteScope);
        const stage3Record = resolveBillingRecordSite(
          record,
          billingRecordSiteCandidates,
        );

        const latestLifecycleRecord = billingRecords
          .filter((candidate) => {
            const candidateScreenCode = normalizeValue(
              candidate.screenCode || candidate.siteScope,
            );
            const sameScreen =
              screenCode && candidateScreenCode === screenCode;
            const sameSource =
              sourceRecordId &&
              normalizeValue(candidate.sourceRecordId) === sourceRecordId;

            return sameScreen || sameSource;
          })
          .sort((left, right) => {
            const leftTime =
              new Date(
                left?.lastUpdatedAt ||
                  left?.submittedAt ||
                  left?.createdAt ||
                  0,
              ).getTime() || 0;
            const rightTime =
              new Date(
                right?.lastUpdatedAt ||
                  right?.submittedAt ||
                  right?.createdAt ||
                  0,
              ).getTime() || 0;

            return rightTime - leftTime;
          })[0];

        const reason = normalizeValue(
          latestLifecycleRecord?.billingStatusReason ||
            stage3Record?.billingLifecycleReason ||
            record.billingStatusReason,
        );
        const billingStatus =
          normalizeValue(
            latestLifecycleRecord?.billingLifecycleStatus ||
              record.billingLifecycleStatus ||
              stage3Record?.billingLifecycleStatus,
          ) || (reason ? "Inactive" : "Active");
        const inactiveSince = normalizeValue(
          stage3Record?.inactiveSince || record.inactiveSince,
        );
        const inactiveSinceMs = inactiveSince
          ? new Date(inactiveSince).getTime()
          : 0;
        const reasonPendingOverdue =
          billingStatus === "Inactive" &&
          !reason &&
          inactiveSinceMs > 0 &&
          nowMs - inactiveSinceMs >= sixHoursMs;

        const firstTimeBillingPending = stage3Record
          ? isPendingFirstTimeBillingRecord(stage3Record)
          : false;

        return {
          ...record,
          billingLifecycleStatus: billingStatus,
          billingStatusReason: reason,
          displayBillingStatus: reason || billingStatus,
          inactiveSince,
          reasonPendingOverdue,
          firstTimeBillingPending,
        };
      })
      .sort((left, right) => {
        if (left.reasonPendingOverdue !== right.reasonPendingOverdue) {
          return left.reasonPendingOverdue ? -1 : 1;
        }

        return 0;
      });
  }, [billingRecordSiteCandidates, billingRecords, canonicalBillingRecordRows]);

  const filteredSubmittedBillingRecords = useMemo(() => {
    const from = billingRecordsFromDate ? parseLocalDateValue(billingRecordsFromDate) : null;
    const to = billingRecordsToDate ? parseLocalDateValue(billingRecordsToDate) : null;

    return submittedBillingRecords
      .filter((record) => {
        if (normalizedBillingRecordsSearch) {
          const matches = [
            record.billingCode,
            record.complexCode,
            record.screenCode,
            record.siteScope,
            record.screenName,
            record.location,
            record.displayBillingStatus,
          ]
            .map((value) => normalizeValue(value).toLowerCase())
            .some((value) => value.includes(normalizedBillingRecordsSearch));
          if (!matches) return false;
        }

        if (from || to) {
          const value =
            normalizeValue(record.submittedAt) ||
            normalizeValue(record.billingPeriodFrom) ||
            normalizeValue(record.billingStartDate) ||
            normalizeValue(record.invoiceDate);
          const date = parseLocalDateValue(String(value).slice(0, 10));
          if (!date) return false;
          if (from && date < from) return false;
          if (to && date > to) return false;
        }
        return true;
      })
      .sort((left, right) => {
        const valueFor = (record) => {
          if (!billingRecordsSortKey) return record.lastUpdatedAt || record.submittedAt || record.billingPeriodFrom || record.billingStartDate || record.invoiceDate || "";
          if (billingRecordsSortKey === "screenCode") return record.screenCode || record.siteScope;
          if (billingRecordsSortKey === "effectiveScreenFee") {
            return getBillingRecordEffectiveScreenFee(
              record,
              commonComplexAllocations,
            );
          }
          if (billingRecordsSortKey === "status") return record.reasonPendingOverdue ? "Inactive — Reason Pending" : record.displayBillingStatus || "Active";
          return record?.[billingRecordsSortKey];
        };
        const diff = compareColumnValues(valueFor(left), valueFor(right));
        return billingRecordsSortOrder === "asc" ? diff : -diff;
      });
  }, [normalizedBillingRecordsSearch, submittedBillingRecords, billingRecordsFromDate, billingRecordsToDate, billingRecordsSortOrder, billingRecordsSortKey, commonComplexAllocations]);

  const billingRecordsCommonPricingSummaries = useMemo(() => {
    const summaries = new Map();

    filteredSubmittedBillingRecords.forEach((billingRecord) => {
      if (!isCommonComplexBillingRecord(billingRecord)) {
        return;
      }

      const groupKey = getCommonComplexBillingGroupKey(billingRecord);
      if (!groupKey || summaries.has(groupKey)) {
        return;
      }

      const liveRepresentative = stage3Records.find(
        (candidate) =>
          isCommonComplexBillingRecord(candidate) &&
          getCommonComplexBillingGroupKey(candidate) === groupKey,
      );
      const summary = getCommonComplexPricingSummary(
        liveRepresentative || billingRecord,
        stage3Records,
        commonComplexAllocations,
        stage2Records,
      );

      if (summary) {
        summaries.set(groupKey, summary);
      }
    });

    return Array.from(summaries.values());
  }, [
    filteredSubmittedBillingRecords,
    stage2Records,
    stage3Records,
    commonComplexAllocations,
  ]);

  const billingRecordsPageCount = Math.max(
    1,
    Math.ceil(
      filteredSubmittedBillingRecords.length / billingRecordsRowsPerPage,
    ),
  );
  const pagedSubmittedBillingRecords = useMemo(() => {
    const startIndex = (billingRecordsPage - 1) * billingRecordsRowsPerPage;
    return filteredSubmittedBillingRecords.slice(
      startIndex,
      startIndex + billingRecordsRowsPerPage,
    );
  }, [
    filteredSubmittedBillingRecords,
    billingRecordsPage,
    billingRecordsRowsPerPage,
  ]);
  const activeBillingRecord = useMemo(
    () =>
      submittedBillingRecords.find(
        (record) => record.billingRecordId === activeBillingRecordId,
      ) || null,
    [activeBillingRecordId, submittedBillingRecords],
  );

  const billingRecordFinancialRows = useMemo(() => {
    if (!activeBillingRecord) {
      return [];
    }

    const targetScreenCode = normalizeValue(
      activeBillingRecord.screenCode || activeBillingRecord.siteScope,
    );

    if (!targetScreenCode) {
      return [];
    }

    const relatedBillingRecords = dedupeBillingRecords(
      billingRecords.filter(
        (record) =>
          normalizeValue(record.screenCode || record.siteScope) ===
          targetScreenCode,
      ),
    );

    const activeStageRecord =
      stage3Records.find(
        (record) =>
          normalizeValue(record.screenCode) === targetScreenCode,
      ) || null;

    const sourceRecord =
      activeStageRecord ||
      relatedBillingRecords.find((record) =>
        normalizeValue(record.billingStartDate),
      ) ||
      activeBillingRecord;

    /*
     * Billing-period lifecycle must survive Reactivate.
     *
     * Reactivation intentionally clears the current "inactive" fields because
     * the site is Active again. The historical pause/inactive event, however,
     * must still cut the old monthly billing periods. Build one lifecycle view
     * for this screen from every matching Stage 3 / Billing record and use the
     * saved closureHistory rather than only the current status fields.
     */
    const lifecycleHistoryById = new Map();

    const lifecycleSourceRecords = [
      ...stage3Records.filter(
        (record) =>
          normalizeValue(record.screenCode) === targetScreenCode,
      ),
      ...relatedBillingRecords,
      activeBillingRecord,
    ].filter(Boolean);

    lifecycleSourceRecords.forEach((record) => {
      const history = Array.isArray(record?.closureHistory)
        ? record.closureHistory
        : [];

      history.forEach((event, index) => {
        const eventKey =
          normalizeValue(event?.eventId) ||
          [
            normalizeValue(event?.newStatus),
            normalizeValue(event?.effectiveDate),
            normalizeValue(event?.updatedAt),
            index,
          ].join("::");

        lifecycleHistoryById.set(eventKey, event);
      });

      /*
       * Backward-compatible fallback for records saved before closureHistory
       * existed. Do not add these when equivalent history is already present.
       */
      const currentClosureDate = normalizeValue(
        record?.pauseFromDate || record?.closureEffectiveDate,
      );
      const currentClosureReason = normalizeValue(
        record?.billingLifecycleReason || record?.billingStatusReason,
      );

      if (currentClosureDate && currentClosureReason) {
        const fallbackKey = `fallback-close::${currentClosureReason}::${currentClosureDate}`;
        if (!lifecycleHistoryById.has(fallbackKey)) {
          lifecycleHistoryById.set(fallbackKey, {
            eventId: fallbackKey,
            previousStatus: "Active",
            newStatus: currentClosureReason,
            effectiveDate: currentClosureDate,
            remarks: normalizeValue(record?.closureRemarks),
            updatedAt: normalizeValue(record?.closureUpdatedAt),
          });
        }
      }

      const currentActiveFrom = normalizeValue(record?.billingActiveFromDate);

      if (currentActiveFrom) {
        const fallbackKey = `fallback-active::${currentActiveFrom}`;
        if (!lifecycleHistoryById.has(fallbackKey)) {
          lifecycleHistoryById.set(fallbackKey, {
            eventId: fallbackKey,
            previousStatus: currentClosureReason || "Inactive",
            newStatus: "Active",
            effectiveDate: currentActiveFrom,
            remarks: "",
            updatedAt: normalizeValue(record?.lastUpdatedAt),
          });
        }
      }
    });

    const lifecycleRecord = {
      ...sourceRecord,
      // Force the period resolver to use the complete history below. Current
      // lifecycle fields can be blank after Reactivate and must not erase an
      // earlier inactive interval.
      billingActiveFromDate: "",
      pauseFromDate: "",
      closureEffectiveDate: "",
      closureHistory: Array.from(lifecycleHistoryById.values()),
    };

    const billingStartDate = parseLocalDateValue(
      sourceRecord?.billingStartDate ||
        activeBillingRecord?.billingStartDate ||
        activeBillingRecord?.billingPeriodFrom,
    );

    if (!billingStartDate) {
      return relatedBillingRecords.flatMap((record) =>
        buildBillingInvoicePeriods(record).map((entry) => {
          const periodDate = parseLocalDateValue(entry.periodFrom);

          return {
            ...entry,
            rowKey: `${record.billingRecordId || getBillingTransactionKey(record)}::${
              entry.entryId
            }`,
            sourceBillingRecordId: record.billingRecordId,
            billingCode: normalizeValue(record.billingCode),
            complexCode: normalizeValue(record.complexCode),
            screenCode: normalizeValue(record.screenCode || record.siteScope),
            screenName: normalizeValue(record.screenName),
            location: normalizeValue(record.location),
            financialYear: getFinancialYearLabelFromDate(periodDate),
            billingAmount:
              normalizeValue(entry.invoiceAmountBeforeGST) ||
              normalizeValue(record.billingAmountBeforeGST),
          };
        }),
      );
    }

    const normalizedMode = normalizeValue(sourceRecord?.subscriptionMode)
      .toLowerCase()
      .replace(/[-_]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    const isMonthly = ["", "monthly", "month"].includes(normalizedMode);

    if (!isMonthly) {
      return relatedBillingRecords.flatMap((record) =>
        buildBillingInvoicePeriods(record).map((entry) => {
          const periodDate = parseLocalDateValue(entry.periodFrom);

          return {
            ...entry,
            rowKey: `${record.billingRecordId || getBillingTransactionKey(record)}::${
              entry.entryId
            }`,
            sourceBillingRecordId: record.billingRecordId,
            billingCode: normalizeValue(record.billingCode),
            complexCode: normalizeValue(record.complexCode),
            screenCode: normalizeValue(record.screenCode || record.siteScope),
            screenName: normalizeValue(record.screenName),
            location: normalizeValue(record.location),
            financialYear: getFinancialYearLabelFromDate(periodDate),
            billingAmount:
              normalizeValue(entry.invoiceAmountBeforeGST) ||
              normalizeValue(record.billingAmountBeforeGST),
          };
        }),
      );
    }

    const overlayByMonth = new Map();

    relatedBillingRecords.forEach((record) => {
      buildBillingInvoicePeriods(record).forEach((entry) => {
        const entryDate = parseLocalDateValue(entry.periodFrom);

        if (!entryDate) {
          return;
        }

        const monthKey = getBillingMonthKey(entryDate);
        const existingOverlay = overlayByMonth.get(monthKey) || {};

        overlayByMonth.set(monthKey, {
          ...existingOverlay,
          ...entry,
          sourceBillingRecordId:
            entry.sourceBillingRecordId ||
            record.billingRecordId ||
            existingOverlay.sourceBillingRecordId ||
            "",
          billingAmount:
            normalizeValue(entry.invoiceAmountBeforeGST) ||
            normalizeValue(record.billingAmountBeforeGST) ||
            normalizeValue(existingOverlay.billingAmount),
          invoiceNumber:
            normalizeValue(entry.invoiceNumber) ||
            normalizeValue(existingOverlay.invoiceNumber),
          invoiceDate:
            normalizeValue(entry.invoiceDate) ||
            normalizeValue(existingOverlay.invoiceDate),
          invoiceAmountBeforeGST:
            normalizeValue(entry.invoiceAmountBeforeGST) ||
            normalizeValue(existingOverlay.invoiceAmountBeforeGST),
          receiptNumber:
            normalizeValue(entry.receiptNumber) ||
            normalizeValue(existingOverlay.receiptNumber),
          receiptDate:
            normalizeValue(entry.receiptDate) ||
            normalizeValue(existingOverlay.receiptDate),
          receivedAmount:
            normalizeValue(entry.receivedAmount) ||
            normalizeValue(existingOverlay.receivedAmount),
          paymentStatus:
            normalizeValue(entry.paymentStatus) ||
            normalizeValue(existingOverlay.paymentStatus),
          paymentRemarks:
            normalizeValue(entry.paymentRemarks) ||
            normalizeValue(existingOverlay.paymentRemarks),
          treatment:
            normalizeValue(entry.treatment) ||
            normalizeValue(existingOverlay.treatment),
          billableAmount:
            normalizeValue(entry.billableAmount) ||
            normalizeValue(existingOverlay.billableAmount),
          invoiceEligibilityLabel:
            normalizeValue(entry.invoiceEligibilityLabel) ||
            normalizeValue(existingOverlay.invoiceEligibilityLabel),
          submittedPeriodTreatment:
            normalizeValue(entry.submittedPeriodTreatment) ||
            normalizeValue(existingOverlay.submittedPeriodTreatment),
          submittedFirstTimePeriod:
            entry.submittedFirstTimePeriod ||
            existingOverlay.submittedFirstTimePeriod ||
            false,
        });
      });
    });

    const today = new Date();
    const currentMonthStart = new Date(
      today.getFullYear(),
      today.getMonth(),
      1,
    );
    let cursor = new Date(
      billingStartDate.getFullYear(),
      billingStartDate.getMonth(),
      1,
    );

    const rows = [];

    while (cursor <= currentMonthStart) {
      const year = cursor.getFullYear();
      const month = cursor.getMonth();
      const monthKey = getBillingMonthKey(cursor);
      const monthStart = new Date(year, month, 1);
      const monthEnd = new Date(year, month + 1, 0);
      const baseEffectiveFrom =
        year === billingStartDate.getFullYear() &&
        month === billingStartDate.getMonth()
          ? billingStartDate
          : monthStart;
      const finalPeriodException = isFinalLifecycleBillingPeriod(
        lifecycleRecord,
        formatLocalDateValue(baseEffectiveFrom),
        formatLocalDateValue(monthEnd),
        today,
      );

      if (
        !isNormalBillingMonthAvailable(cursor, today) &&
        !finalPeriodException
      ) {
        cursor = new Date(year, month + 1, 1);
        continue;
      }

      const lifecyclePeriod = getLifecycleAdjustedMonthlyPeriod(
        lifecycleRecord,
        formatLocalDateValue(baseEffectiveFrom),
        formatLocalDateValue(monthEnd),
      );
      const effectiveFrom = lifecyclePeriod.billable
        ? parseLocalDateValue(lifecyclePeriod.periodFrom)
        : null;
      const effectiveTo = lifecyclePeriod.billable
        ? parseLocalDateValue(lifecyclePeriod.periodTo)
        : null;
      const overlay = overlayByMonth.get(monthKey) || {};
      const isSubmittedWaivedPeriod =
        normalizeValue(overlay.submittedPeriodTreatment).toLowerCase() ===
        "waived";

      rows.push({
        ...overlay,
        entryId: monthKey,
        rowKey: `${targetScreenCode}::${monthKey}`,
        billingMonth: getBillingMonthLabel(cursor),
        periodFrom: lifecyclePeriod.billable
          ? formatLocalDateValue(effectiveFrom)
          : "",
        periodTo: lifecyclePeriod.billable
          ? formatLocalDateValue(effectiveTo)
          : "",
        billingCode: normalizeValue(sourceRecord?.billingCode),
        complexCode: normalizeValue(sourceRecord?.complexCode),
        screenCode: targetScreenCode,
        screenName: normalizeValue(sourceRecord?.screenName),
        location: normalizeValue(sourceRecord?.location),
        financialYear: getFinancialYearLabelFromDate(cursor),
        billingAmount:
          normalizeValue(overlay.billingAmount) ||
          normalizeValue(overlay.invoiceAmountBeforeGST) ||
          "",
        invoiceNumber: normalizeValue(overlay.invoiceNumber),
        invoiceDate: normalizeValue(overlay.invoiceDate),
        invoiceAmountBeforeGST: normalizeValue(
          overlay.invoiceAmountBeforeGST,
        ),
        receiptNumber: normalizeValue(overlay.receiptNumber),
        receiptDate: normalizeValue(overlay.receiptDate),
        receivedAmount: normalizeValue(overlay.receivedAmount),
        paymentStatus: lifecyclePeriod.billable
          ? isSubmittedWaivedPeriod
            ? "Waived"
            : normalizeValue(overlay.paymentStatus) ||
            (normalizeValue(overlay.invoiceNumber) ? "Pending" : "Not Invoiced")
          : "Paused / Not Billable",
        paymentRemarks: lifecyclePeriod.billable
          ? normalizeValue(overlay.paymentRemarks)
          : lifecyclePeriod.lifecycleRemarks || normalizeValue(overlay.paymentRemarks),
        lifecycleBillable: lifecyclePeriod.billable,
        lifecycleStatus: lifecyclePeriod.lifecycleStatus || (lifecyclePeriod.billable ? "Billable" : "Paused / Not Billable"),
        lifecycleRemarks: lifecyclePeriod.lifecycleRemarks || "",
        lifecyclePeriodLabel: lifecyclePeriod.lifecyclePeriodLabel || "",
        submittedPeriodTreatment: normalizeValue(
          overlay.submittedPeriodTreatment,
        ),
        submittedFirstTimePeriod: Boolean(overlay.submittedFirstTimePeriod),
      });

      cursor = new Date(year, month + 1, 1);
    }

    return rows.sort((left, right) => {
      const leftDate = parseLocalDateValue(left.periodFrom);
      const rightDate = parseLocalDateValue(right.periodFrom);
      return (rightDate?.getTime() || 0) - (leftDate?.getTime() || 0);
    });
  }, [activeBillingRecord, billingRecords, stage3Records]);

  const billingRecordFinancialYearOptions = useMemo(() => {
    const years = Array.from(
      new Set(
        billingRecordFinancialRows
          .map((row) => row.financialYear)
          .filter(Boolean),
      ),
    ).sort((left, right) => right.localeCompare(left));

    return years;
  }, [billingRecordFinancialRows]);

  const billingRecordMonthOptions = useMemo(() => {
    const rowsForYear =
      billingRecordFinancialYear === "All"
        ? billingRecordFinancialRows
        : billingRecordFinancialRows.filter(
            (row) => row.financialYear === billingRecordFinancialYear,
          );

    return rowsForYear.map((row) => ({
      value: row.entryId,
      label: row.billingMonth,
    }));
  }, [billingRecordFinancialRows, billingRecordFinancialYear]);

  const filteredBillingRecordFinancialRows = useMemo(() => {
    const rowsForYear =
      billingRecordFinancialYear === "All"
        ? billingRecordFinancialRows
        : billingRecordFinancialRows.filter(
            (row) => row.financialYear === billingRecordFinancialYear,
          );

    if (billingRecordMonthFilter === "All") {
      return rowsForYear;
    }

    return rowsForYear.filter(
      (row) => row.entryId === billingRecordMonthFilter,
    );
  }, [
    billingRecordFinancialRows,
    billingRecordFinancialYear,
    billingRecordMonthFilter,
  ]);

  const selectedBillingRecordFinancialRow = useMemo(
    () =>
      billingRecordFinancialRows.find(
        (row) => row.rowKey === billingRecordSelectedEntryKey,
      ) || null,
    [billingRecordFinancialRows, billingRecordSelectedEntryKey],
  );
  const isSelectedBillingMonthLifecycleLocked =
    selectedBillingRecordFinancialRow?.lifecycleBillable === false;
  const isSelectedBillingMonthWaived =
    normalizeValue(
      selectedBillingRecordFinancialRow?.submittedPeriodTreatment,
    ).toLowerCase() === "waived";
  const selectedBillingRecordPaymentDate = normalizeValue(
    selectedBillingRecordFinancialRow?.paymentReceivedDate ||
      selectedBillingRecordFinancialRow?.receiptDate,
  );
  const selectedBillingRecordHasFuturePayment = Boolean(
    selectedBillingRecordFinancialRow &&
      isPaymentReceivedStatus(selectedBillingRecordFinancialRow.paymentStatus) &&
      isFutureLocalDateValue(selectedBillingRecordPaymentDate),
  );

  const billingHistoryRecords = useMemo(() => {
    const targetScreenCode = normalizeValue(billingHistoryScreenCode);

    if (!targetScreenCode) {
      return [];
    }

    return dedupeBillingRecords(
      billingRecords.filter(
        (record) =>
          normalizeValue(record.screenCode || record.siteScope) ===
          targetScreenCode,
      ),
    )
      .slice()
      .sort((left, right) => {
        const leftTime = new Date(
          left.submittedAt || left.createdAt || left.lastUpdatedAt || 0,
        ).getTime();
        const rightTime = new Date(
          right.submittedAt || right.createdAt || right.lastUpdatedAt || 0,
        ).getTime();

        return rightTime - leftTime;
      });
  }, [billingHistoryScreenCode, billingRecords]);

  const billingHistorySite = useMemo(() => {
    if (!billingHistoryScreenCode) {
      return null;
    }

    return (
      stage3Records.find(
        (record) =>
          normalizeValue(record.screenCode) ===
          normalizeValue(billingHistoryScreenCode),
      ) ||
      billingHistoryRecords[0] ||
      null
    );
  }, [billingHistoryScreenCode, billingHistoryRecords, stage3Records]);

  const billingHistoryAuditEntries = useMemo(() => {
    const entries = [];

    billingHistoryRecords.forEach((record) => {
      const submittedAt = record.submittedAt || record.createdAt;

      if (submittedAt) {
        entries.push({
          auditId: `submitted-${record.billingRecordId || submittedAt}`,
          timestamp: submittedAt,
          event: "Billing Submitted",
          field: "billingVerificationStatus",
          previousValue: "-",
          newValue: record.submissionStatus || "Submitted to Billing Team",
          remarks: formatBillingPeriod(
            record.billingPeriodFrom,
            record.billingPeriodTo,
          ),
        });
      }

      const auditTrail = Array.isArray(record.auditTrail)
        ? record.auditTrail
        : [];

      auditTrail.forEach((entry, index) => {
        entries.push({
          auditId:
            entry.auditId ||
            `${record.billingRecordId || "billing-record"}-audit-${index}`,
          timestamp:
            entry.timestamp ||
            entry.changedAt ||
            entry.createdAt ||
            record.lastUpdatedAt ||
            record.submittedAt ||
            "",
          event: entry.event || entry.action || "Updated",
          field: entry.field || "",
          previousValue: formatAuditValue(entry.previousValue),
          newValue: formatAuditValue(entry.newValue),
          remarks: entry.remarks || entry.reason || "",
        });
      });
    });

    return entries.sort((left, right) => {
      const leftTime = new Date(left.timestamp || 0).getTime() || 0;
      const rightTime = new Date(right.timestamp || 0).getTime() || 0;
      return rightTime - leftTime;
    });
  }, [billingHistoryRecords]);

  const combinedPaymentRows = useMemo(
    () => buildCombinedPaymentDownloadRows(otfTransactions, stage3Records),
    [otfTransactions, stage3Records],
  );
  const hasCombinedPaymentRows = combinedPaymentRows.length > 0;
  const isOtfApplicable = normalizeValue(activeRecord?.otfApplicable) === "Yes";
  const activeValidationErrors = useMemo(
    () => getBillingValidationErrors(activeRecord),
    [activeRecord],
  );

  const activeSiteCommonComplexBillingRow = useMemo(() => {
    if (!activeRecord || !isCommonComplexBillingRecord(activeRecord)) {
      return null;
    }

    return (
      commonComplexBillingRows.find(
        (row) =>
          row.siteKey === getCommonComplexBillingSiteKey(activeRecord.screenCode),
      ) || null
    );
  }, [activeRecord, commonComplexBillingRows]);

  const activeSavedSiteBillingAllocation = useMemo(
    () =>
      activeRecord
        ? getSavedSiteBillingAllocation(activeRecord, commonComplexAllocations)
        : "",
    [activeRecord, commonComplexAllocations],
  );

  const activeSiteWiseSubscriptionFee = useMemo(
    () => (activeRecord ? resolveSiteWiseSubscriptionFee(activeRecord) : ""),
    [activeRecord],
  );

  const activeCommonComplexAllocationComplete = useMemo(() => {
    if (!activeRecord || !isCommonComplexBillingRecord(activeRecord)) {
      return false;
    }

    return (
      getCommonComplexBillingValidationError(
        activeRecord,
        stage3Records,
        commonComplexAllocations,
      ) === ""
    );
  }, [activeRecord, stage3Records, commonComplexAllocations]);

  const monthlyBillingRows = useMemo(
    () =>
      activeRecord
        ? buildMonthlyBillingRows(
            activeRecord,
            isVariableSubscriptionRecord(activeRecord)
              ? undefined
              : isCommonComplexBillingRecord(activeRecord)
                ? activeCommonComplexAllocationComplete
                  ? activeSiteCommonComplexBillingRow?.allocatedFee ||
                      activeSavedSiteBillingAllocation ||
                      ""
                  : ""
                : activeSiteWiseSubscriptionFee ||
                    normalizeValue(activeRecord.subscriptionFee),
          )
        : [],
    [
      activeRecord,
      activeCommonComplexAllocationComplete,
      activeSiteCommonComplexBillingRow,
      activeSavedSiteBillingAllocation,
      activeSiteWiseSubscriptionFee,
    ],
  );

  const firstTimeMonthlyBillingRows = useMemo(() => {
    const completedPeriodKeys = getCompletedMonthlyPeriodKeys(
      activeRecord,
      billingRecords,
    );

    return monthlyBillingRows.filter(
      (row) => !completedPeriodKeys.has(row.monthKey),
    );
  }, [activeRecord, billingRecords, monthlyBillingRows]);

  const completedActivePeriodKeys = useMemo(
    () => getCompletedMonthlyPeriodKeys(activeRecord, billingRecords),
    [activeRecord, billingRecords],
  );

  const firstTimePeriodBreakdown = useMemo(
    () => buildFirstTimePeriodBreakdown(firstTimeMonthlyBillingRows, firstTimeWaiverDraft),
    [firstTimeMonthlyBillingRows, firstTimeWaiverDraft],
  );

  const firstTimeBillingSummary = useMemo(() => {
    const totalBeforeWaiver = firstTimePeriodBreakdown.reduce(
      (total, row) => total + (parseAmountValue(row.normalAmount) || 0),
      0,
    );
    const actualBillingAmount = firstTimePeriodBreakdown.reduce(
      (total, row) => total + (parseAmountValue(row.billableAmount) || 0),
      0,
    );

    return {
      totalBeforeWaiver: formatAmountValue(totalBeforeWaiver),
      waiverAmount: formatAmountValue(totalBeforeWaiver - actualBillingAmount),
      actualBillingAmount: formatAmountValue(actualBillingAmount),
    };
  }, [firstTimePeriodBreakdown]);

  useEffect(() => {
    const savedWaiver = activeRecord?.firstTimeWaiver;
    setFirstTimeWaiverDraft({
      applicable: savedWaiver?.applicable === "Yes" ? "Yes" : "No",
      from: normalizeValue(savedWaiver?.from),
      to: normalizeValue(savedWaiver?.to),
      reason: normalizeValue(savedWaiver?.reason),
    });
  }, [activeRecordId]);

  const normalizedActiveSubscriptionMode = normalizeValue(
    activeRecord?.subscriptionMode,
  )
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const isMonthlyBilling =
    normalizedActiveSubscriptionMode === "monthly" ||
    normalizedActiveSubscriptionMode === "month" ||
    normalizedActiveSubscriptionMode === "";


  const focRows = useMemo(() => {
    const query = normalizeValue(focSearchTerm).toLowerCase();
    const from = focFromDate ? parseLocalDateValue(focFromDate) : null;
    const to = focToDate ? parseLocalDateValue(focToDate) : null;

    return stage3Records
      .filter((record) => isFoCOriginRecord(record) && isFoCBillingDateEligible(record))
      .filter((record) => {
        if (query) {
          const matches = [
            record.billingCode,
            record.complexCode,
            record.billingName,
            record.screenCode,
            record.erpScreenName,
            record.screenName,
            record.location,
            record.state,
            getFoCCurrentCommercialStatus(record),
          ]
            .map((value) => normalizeValue(value).toLowerCase())
            .some((value) => value.includes(query));
          if (!matches) return false;
        }
        if (from || to) {
          const date = parseLocalDateValue(normalizeValue(record.billingStartDate || record.billingDate));
          if (!date) return false;
          if (from && date < from) return false;
          if (to && date > to) return false;
        }
        return true;
      })
      .sort((left, right) => {
        const valueFor = (record) => {
          if (!focSortKey || focSortKey === "billingDate") return record.billingStartDate || record.billingDate;
          if (focSortKey === "commercialStatus") return getFoCCurrentCommercialStatus(record);
          if (focSortKey === "erpScreenName") return record.erpScreenName || record.screenName;
          if (focSortKey === "otf") return normalizeValue(record.focBillableOtfAmount) || "0";
          if (focSortKey === "subscription") return normalizeValue(record.focBillableSubscriptionFee) || "0";
          return record?.[focSortKey];
        };
        const diff = compareColumnValues(valueFor(left), valueFor(right));
        return focSortOrder === "asc" ? diff : -diff;
      });
  }, [stage3Records, focSearchTerm, focFromDate, focToDate, focSortOrder, focSortKey]);

  const focPageCount = Math.max(1, Math.ceil(focRows.length / focRowsPerPage));
  const pagedFocRows = useMemo(() => {
    const startIndex = (focPage - 1) * focRowsPerPage;
    return focRows.slice(startIndex, startIndex + focRowsPerPage);
  }, [focRows, focPage, focRowsPerPage]);

  useEffect(() => {
    setFocPage(1);
  }, [focSearchTerm, focFromDate, focToDate, focSortKey, focSortOrder, focRowsPerPage]);

  useEffect(() => {
    setFocPage((currentPage) => Math.min(currentPage, focPageCount));
  }, [focPageCount]);

  const activeFocRecord = useMemo(
    () => focRows.find((record) => record.recordId === activeFocRecordId) || null,
    [focRows, activeFocRecordId],
  );

  function handleOpenFocBillableEditor(record) {
    if (!record) return;

    setActiveFocRecordId(record.recordId);
    setFocSaveMessage("");
    setFocBillableDraft({
      effectiveFrom: normalizeValue(record.focBillableEffectiveFrom),
      otfApplicable: normalizeValue(record.focBillableOtfApplicable) || "No",
      otfAmount: normalizeValue(record.focBillableOtfAmount),
      subscriptionApplicable:
        normalizeValue(record.focBillableSubscriptionApplicable) || "Yes",
      subscriptionFee: normalizeValue(record.focBillableSubscriptionFee),
      remarks: "",
    });
    setIsFocBillableEditorOpen(true);
  }

  function handleCloseFocBillableEditor() {
    setIsFocBillableEditorOpen(false);
    setActiveFocRecordId("");
    setFocSaveMessage("");
  }

  function updateFocBillableDraft(field, value) {
    setFocBillableDraft((current) => ({ ...current, [field]: value }));
    setFocSaveMessage("");
  }

  function handleSaveFocBillableChange() {
    if (!activeFocRecord) return;

    const effectiveFrom = normalizeValue(focBillableDraft.effectiveFrom);
    const remarks = normalizeValue(focBillableDraft.remarks);
    const otfApplicable = normalizeValue(focBillableDraft.otfApplicable) || "No";
    const subscriptionApplicable =
      normalizeValue(focBillableDraft.subscriptionApplicable) || "No";
    const otfAmount = sanitizeAmount(focBillableDraft.otfAmount);
    const subscriptionFee = sanitizeAmount(focBillableDraft.subscriptionFee);

    if (!effectiveFrom || !isValidDateValue(effectiveFrom)) {
      setFocSaveMessage("A valid Effective From date is required.");
      return;
    }

    if (!remarks) {
      setFocSaveMessage("Remarks are required for FoC commercial change.");
      return;
    }

    if (otfApplicable === "Yes" && !otfAmount) {
      setFocSaveMessage("Enter OTF Amount when OTF is Applicable.");
      return;
    }

    if (subscriptionApplicable === "Yes" && !subscriptionFee) {
      setFocSaveMessage("Enter Subscription Fee when Subscription is Applicable.");
      return;
    }

    const changedAt = new Date().toISOString();

    setStage3Records((currentRecords) =>
      currentRecords.map((record) => {
        if (record.recordId !== activeFocRecord.recordId) return record;

        const previousStatus = getFoCCurrentCommercialStatus(record);
        const history = Array.isArray(record.focCommercialHistory)
          ? record.focCommercialHistory
          : [];

        return {
          ...record,
          focOrigin: true,
          originalCommercialStatus: "FoC",
          focCurrentCommercialStatus: "Billable",
          focBillableEffectiveFrom: effectiveFrom,
          focBillableOtfApplicable: otfApplicable,
          focBillableOtfAmount: otfApplicable === "Yes" ? otfAmount : "0",
          focBillableSubscriptionApplicable: subscriptionApplicable,
          focBillableSubscriptionFee:
            subscriptionApplicable === "Yes" ? subscriptionFee : "0",
          focCommercialUpdatedAt: changedAt,
          focCommercialHistory: [
            ...history,
            {
              historyId: `foc-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
              previousStatus,
              newStatus: "Billable",
              effectiveFrom,
              otfApplicable,
              otfAmount: otfApplicable === "Yes" ? otfAmount : "0",
              subscriptionApplicable,
              subscriptionFee:
                subscriptionApplicable === "Yes" ? subscriptionFee : "0",
              remarks,
              changedAt,
            },
          ],
        };
      }),
    );

    setFocSaveMessage("FoC commercial status changed to Billable. The site remains in FoC.");
    setIsFocBillableEditorOpen(false);
  }

  const closureRows = useMemo(() => {
    const merged = new Map();

    stage3Records.forEach((record) => {
      const reason = normalizeValue(record.billingLifecycleReason);
      const billingStatus =
        normalizeValue(record.billingLifecycleStatus) ||
        (reason ? "Inactive" : "Active");

      if (billingStatus !== "Inactive") {
        return;
      }

      const pendingFirstTimeBilling =
        isPendingFirstTimeBillingRecord(record);

      if (pendingFirstTimeBilling) {
        return;
      }

      const screenCode = normalizeValue(record.screenCode);
      const key =
        screenCode ||
        normalizeValue(record.recordId) ||
        `${normalizeValue(record.billingCode)}::${normalizeValue(record.screenName)}`;

      merged.set(key, {
        ...record,
        lifecycleSourceRecordId: normalizeValue(record.recordId),
        billingLifecycleStatus: "Inactive",
        currentBillingStatus: reason || "Inactive",
      });
    });

    billingRecords.forEach((record) => {
      const reason = normalizeValue(
        record.billingStatusReason || record.billingLifecycleReason,
      );
      const billingStatus =
        normalizeValue(record.billingLifecycleStatus) ||
        (reason ? "Inactive" : "Active");

      if (billingStatus !== "Inactive") {
        return;
      }

      const matchingStage3Record =
        stage3Records.find(
          (site) =>
            normalizeValue(record.sourceRecordId) &&
            normalizeValue(site.recordId) ===
              normalizeValue(record.sourceRecordId),
        ) ||
        stage3Records.find(
          (site) =>
            normalizeValue(record.screenCode || record.siteScope) &&
            normalizeValue(site.screenCode) ===
              normalizeValue(record.screenCode || record.siteScope),
        );

      if (
        matchingStage3Record &&
        isPendingFirstTimeBillingRecord(matchingStage3Record)
      ) {
        return;
      }

      const screenCode = normalizeValue(record.screenCode || record.siteScope);
      const key =
        screenCode ||
        normalizeValue(record.sourceRecordId) ||
        normalizeValue(record.billingRecordId) ||
        `${normalizeValue(record.billingCode)}::${normalizeValue(record.screenName)}`;

      const existing = merged.get(key);

      merged.set(key, {
        ...(existing || {}),
        ...record,
        recordId:
          normalizeValue(existing?.recordId) ||
          normalizeValue(record.sourceRecordId) ||
          normalizeValue(record.recordId) ||
          key,
        lifecycleSourceRecordId:
          normalizeValue(existing?.lifecycleSourceRecordId) ||
          normalizeValue(record.sourceRecordId),
        billingRecordId: normalizeValue(record.billingRecordId),
        screenCode,
        billingLifecycleStatus: "Inactive",
        currentBillingStatus: reason || existing?.currentBillingStatus || "Inactive",
      });
    });

    return Array.from(merged.values()).sort((left, right) => {
      const leftPending = left.currentBillingStatus === "Inactive";
      const rightPending = right.currentBillingStatus === "Inactive";

      if (leftPending !== rightPending) {
        return leftPending ? -1 : 1;
      }

      return (left.screenName || left.screenCode || "").localeCompare(
        right.screenName || right.screenCode || "",
      );
    });
  }, [stage3Records, billingRecords]);

  const filteredClosureRows = useMemo(() => {
    const filtered = normalizedClosureSearch
      ? closureRows.filter((record) =>
          [
            record.billingCode,
            record.complexCode,
            record.screenCode,
            record.screenName,
            record.location,
            record.currentBillingStatus,
            record.closureEffectiveDate,
            record.closureRemarks,
          ]
            .map((value) => normalizeValue(value).toLowerCase())
            .some((value) => value.includes(normalizedClosureSearch)),
        )
      : closureRows;

    return [...filtered].sort((left, right) => {
      const valueFor = (record) => closureSortKey ? record?.[closureSortKey] : record.closureEffectiveDate || record.inactiveEffectiveDate || record.pauseFromDate;
      const diff = compareColumnValues(valueFor(left), valueFor(right));
      return closureSortOrder === "asc" ? diff : -diff;
    });
  }, [closureRows, normalizedClosureSearch, closureSortOrder, closureSortKey]);

  const closurePageCount = Math.max(
    1,
    Math.ceil(filteredClosureRows.length / closureRowsPerPage),
  );

  const pagedClosureRows = useMemo(() => {
    const startIndex = (closurePage - 1) * closureRowsPerPage;
    return filteredClosureRows.slice(
      startIndex,
      startIndex + closureRowsPerPage,
    );
  }, [filteredClosureRows, closurePage, closureRowsPerPage]);

  const activeClosureRecord = useMemo(
    () =>
      closureRows.find(
        (record) => record.recordId === activeClosureRecordId,
      ) || null,
    [closureRows, activeClosureRecordId],
  );

  const visibleVerificationRowIds = useMemo(
    () => new Set(selectableVerificationRows.map((record) => record.recordId)),
    [selectableVerificationRows],
  );

  const selectableVerificationRowIds = useMemo(
    () => new Set(selectableVerificationRows.map((record) => record.recordId)),
    [selectableVerificationRows],
  );

  const hasVisibleVerificationRows = firstTimeVisibleRows.length > 0;

  const pagedSelectableVerificationRows = useMemo(
    () =>
      pagedVerificationRows.filter((record) =>
        isFirstTimeBillingSiteReadyForSubmission(
          record,
          stage3Records,
          commonComplexAllocations,
        ),
      ),
    [
      pagedVerificationRows,
      stage3Records,
      commonComplexAllocations,
    ],
  );

  const areAllVisibleVerificationRowsSelected =
    pagedSelectableVerificationRows.length > 0 &&
    pagedSelectableVerificationRows.every((record) =>
      selectedVerificationIds.includes(record.recordId),
    );
  const areSomeVisibleVerificationRowsSelected =
    pagedSelectableVerificationRows.some((record) =>
      selectedVerificationIds.includes(record.recordId),
    ) &&
    !areAllVisibleVerificationRowsSelected;
  const hasSelectedVerificationRows = selectableVerificationRows.some((record) =>
    selectedVerificationIds.includes(record.recordId),
  );

  useEffect(() => {
    setFirstTimePage(1);
  }, [normalizedSearch, firstTimeRowsPerPage, firstTimeStatusFilter]);

  useEffect(() => {
    setRecurringPage(1);
  }, [normalizedRecurringSearch, recurringRowsPerPage]);

  useEffect(() => {
    setBillingRecordsPage(1);
  }, [normalizedBillingRecordsSearch, billingRecordsRowsPerPage]);

  useEffect(() => {
    setClosurePage(1);
  }, [normalizedClosureSearch, closureRowsPerPage]);

  useEffect(() => {
    if (firstTimePage > firstTimePageCount) {
      setFirstTimePage(firstTimePageCount);
    }
  }, [firstTimePage, firstTimePageCount]);

  useEffect(() => {
    if (recurringPage > recurringPageCount) {
      setRecurringPage(recurringPageCount);
    }
  }, [recurringPage, recurringPageCount]);

  useEffect(() => {
    if (billingRecordsPage > billingRecordsPageCount) {
      setBillingRecordsPage(billingRecordsPageCount);
    }
  }, [billingRecordsPage, billingRecordsPageCount]);

  useEffect(() => {
    if (closurePage > closurePageCount) {
      setClosurePage(closurePageCount);
    }
  }, [closurePage, closurePageCount]);

  useEffect(() => {
    if (!verificationSelectAllRef.current) {
      return;
    }

    verificationSelectAllRef.current.indeterminate =
      areSomeVisibleVerificationRowsSelected;
  }, [areSomeVisibleVerificationRowsSelected]);

  useEffect(() => {
    if (!activeOtfTransactionId) return;

    const activeTransactionExists = otfTransactions.some(
      (transaction) => transaction.transactionId === activeOtfTransactionId,
    );

    if (!activeTransactionExists) {
      setActiveOtfTransactionId("");
    }
  }, [activeOtfTransactionId, otfTransactions]);

  useEffect(() => {
    setOtfTransactionEdits((currentEdits) => {
      const nextEdits = {};

      otfTransactions.forEach((transaction) => {
        if (currentEdits[transaction.transactionId]) {
          nextEdits[transaction.transactionId] =
            currentEdits[transaction.transactionId];
        }
      });

      const currentKeys = Object.keys(currentEdits);
      const nextKeys = Object.keys(nextEdits);
      const unchanged =
        currentKeys.length === nextKeys.length &&
        nextKeys.every(
          (key) => currentEdits[key] === nextEdits[key],
        );

      return unchanged ? currentEdits : nextEdits;
    });
  }, [otfTransactions]);

  useEffect(() => {
    if (!activeRecord) {
      return;
    }

    const billingStartDate = normalizeValue(activeRecord.billingStartDate);
    const billingPeriodFrom = normalizeValue(activeRecord.billingPeriodFrom);
    const billingPeriodTo = normalizeValue(activeRecord.billingPeriodTo);

    if (!billingPeriodFrom && billingStartDate) {
      updateActiveRecord("billingPeriodFrom", billingStartDate);
      return;
    }

    if (billingPeriodFrom && isValidDateValue(billingPeriodFrom)) {
      const nextBillingPeriodTo = calculateBillingPeriodTo(
        billingPeriodFrom,
        activeRecord,
      );

      if (nextBillingPeriodTo && nextBillingPeriodTo !== billingPeriodTo) {
        updateActiveRecord("billingPeriodTo", nextBillingPeriodTo);
      }
    }
  }, [
    activeRecord?.recordId,
    activeRecord?.billingStartDate,
    activeRecord?.billingPeriodFrom,
    activeRecord?.billingPeriodTo,
    activeRecord?.subscriptionMode,
    activeRecord,
  ]);

  useEffect(() => {
    if (activeAllocationSiteKey && !activeCommonComplexBillingRow) {
      setActiveAllocationSiteKey("");
      setAllocationDraftFee("");
      setAllocationSaveMessage("");
    }
  }, [activeAllocationSiteKey, activeCommonComplexBillingRow]);

  useEffect(() => {
    if (!activeBillingRecord) {
      setBillingRecordDraft(null);
      setBillingRecordFinancialYear("");
      setBillingRecordMonthFilter("All");
      setBillingRecordSelectedEntryKey("");
      setBillingRecordMonthDetailMode("view");
      return;
    }

    setBillingRecordDraft({
      billingRecordId: activeBillingRecord.billingRecordId,
      billingCode: normalizeValue(activeBillingRecord.billingCode),
      complexCode: normalizeValue(activeBillingRecord.complexCode),
      siteScope: normalizeValue(activeBillingRecord.siteScope),
      billingPeriodFrom: normalizeValue(activeBillingRecord.billingPeriodFrom),
      billingPeriodTo: normalizeValue(activeBillingRecord.billingPeriodTo),
      billingAmountBeforeGST: normalizeValue(
        activeBillingRecord.billingAmountBeforeGST,
      ),
      billingMode: normalizeValue(activeBillingRecord.billingMode),
      submissionStatus: normalizeValue(activeBillingRecord.submissionStatus),
      submittedAt: normalizeValue(activeBillingRecord.submittedAt),
      createdAt: normalizeValue(activeBillingRecord.createdAt),
      lastUpdatedAt: normalizeValue(activeBillingRecord.lastUpdatedAt),
      invoiceNumber: normalizeValue(activeBillingRecord.invoiceNumber),
      invoiceDate: normalizeValue(activeBillingRecord.invoiceDate),
      invoiceAmountBeforeGST: normalizeValue(
        activeBillingRecord.invoiceAmountBeforeGST,
      ),
      invoiceStatus: normalizeValue(activeBillingRecord.invoiceStatus),
      paymentStatus: normalizeValue(activeBillingRecord.paymentStatus) || "Pending",
      paymentReceivedDate: normalizeValue(
        activeBillingRecord.paymentReceivedDate,
      ),
      paymentRemarks: normalizeValue(activeBillingRecord.paymentRemarks),
      billingRemarks: normalizeValue(activeBillingRecord.billingRemarks),
      billingStatus: (() => {
        const sourceRecordId = normalizeValue(activeBillingRecord.sourceRecordId);
        const screenCode = normalizeValue(
          activeBillingRecord.screenCode || activeBillingRecord.siteScope,
        );
        const stage3Record =
          stage3Records.find(
            (site) =>
              sourceRecordId &&
              normalizeValue(site.recordId) === sourceRecordId,
          ) ||
          stage3Records.find(
            (site) =>
              screenCode &&
              normalizeValue(site.screenCode) === screenCode,
          );
        const reason = normalizeValue(stage3Record?.billingLifecycleReason);
        return (
          normalizeValue(activeBillingRecord.billingLifecycleStatus) ||
          normalizeValue(stage3Record?.billingLifecycleStatus) ||
          (reason ? "Inactive" : "Active")
        );
      })(),
      invoiceEntries: buildBillingInvoicePeriods(activeBillingRecord),
    });
  }, [activeBillingRecord, stage3Records]);

  useEffect(() => {
    if (!activeBillingRecord) {
      return;
    }

    const currentFinancialYear = getCurrentFinancialYearLabel();
    const preferredFinancialYear = billingRecordFinancialYearOptions.includes(
      currentFinancialYear,
    )
      ? currentFinancialYear
      : billingRecordFinancialYearOptions[0] || "All";

    setBillingRecordFinancialYear(preferredFinancialYear);
    setBillingRecordMonthFilter("All");
    setBillingRecordSelectedEntryKey("");
    setBillingRecordMonthDetailMode("view");
    setBillingRecordManualPaymentDraft(null);
    setBillingRecordManualPaymentMessage("");
  }, [activeBillingRecord?.billingRecordId, billingRecordFinancialYearOptions]);

  useEffect(() => {
    if (
      activeBillingRecordId &&
      !submittedBillingRecords.some(
        (record) => record.billingRecordId === activeBillingRecordId,
      )
    ) {
      setActiveBillingRecordId("");
      setBillingRecordDraft(null);
      setBillingRecordSaveMessage("");
    }
  }, [activeBillingRecordId, submittedBillingRecords]);

  useEffect(() => {
    setSelectedVerificationIds((currentIds) =>
      currentIds.filter((recordId) => visibleVerificationRowIds.has(recordId)),
    );
  }, [visibleVerificationRowIds]);

  useEffect(() => {
    if (
      activeRecordId &&
      !stage3Records.some((record) => record.recordId === activeRecordId) &&
      !recurringRows.some((record) => record.recordId === activeRecordId)
    ) {
      setActiveRecordId("");
    }
  }, [activeRecordId, stage3Records, recurringRows]);

  function updateActiveRecord(field, value) {
    if (!activeRecord) {
      return;
    }

    setStage3Records((currentRecords) =>
      currentRecords.map((record) => {
        if (record.recordId !== activeRecord.recordId) return record;

        const nextRecord = { ...record, [field]: value };
        if (normalizeValue(record[field]) === normalizeValue(value)) {
          return record;
        }

        return invalidateFirstTimeBillingValidation(nextRecord, field);
      }),
    );
  }

  function handleVariableMonthlyAmountChange(monthKey, value) {
    if (!activeRecord) {
      return;
    }

    const nextAmounts = {
      ...(activeRecord.monthlyBillingAmounts || {}),
      [monthKey]: sanitizeAmount(value),
    };

    setFirstTimeValidationMessage("");

    setStage3Records((currentRecords) =>
      currentRecords.map((record) =>
        record.recordId === activeRecord.recordId
          ? {
              ...invalidateFirstTimeBillingValidation(
                record,
                "monthlyBillingAmounts",
              ),
              monthlyBillingAmounts: nextAmounts,
            }
          : record,
      ),
    );
  }

  function handleValidatedDateChange(field) {
    return (event) => {
      const nextValue = event.target.value;

      updateActiveRecord(field, nextValue);
    };
  }

  function handleBillingPeriodFromChange(event) {
    const nextValue = event.target.value;

    updateActiveRecord("billingPeriodFrom", nextValue);

    if (!isValidDateValue(nextValue)) {
      return;
    }

    const nextBillingPeriodTo = calculateBillingPeriodTo(nextValue, activeRecord);
    if (nextBillingPeriodTo) {
      updateActiveRecord("billingPeriodTo", nextBillingPeriodTo);
    }
  }

  function handleValidatedDateBlur(field) {
    return (event) => {
      const nextValue = event.target.value;

      if (nextValue && !isValidDateValue(nextValue)) {
        alert("Year must contain exactly 4 digits.");
      }
    };
  }

  function updateCommonComplexAllocation(groupKey, siteKey, nextValue) {
    setCommonComplexAllocations((currentAllocations) => {
      const currentGroupAllocations = currentAllocations[groupKey] || {};
      const currentEntry = currentGroupAllocations[siteKey] || {};

      if (currentEntry.locked) {
        return currentAllocations;
      }

      return {
        ...currentAllocations,
        [groupKey]: {
          ...currentGroupAllocations,
          [siteKey]: {
            ...currentEntry,
            allocatedFee: sanitizeAmount(nextValue),
          },
        },
      };
    });
  }

  function lockCommonComplexAllocations(records) {
    if (!records.length) {
      return;
    }

    const submittedGroupKeys = new Set(
      records
        .filter((record) => isCommonComplexBillingRecord(record))
        .map((record) => getCommonComplexBillingGroupKey(record))
        .filter(Boolean),
    );

    if (submittedGroupKeys.size === 0) {
      return;
    }

    const lockedAt = getCurrentLocalTimestamp();

    setCommonComplexAllocations((currentAllocations) => {
      let hasChanges = false;
      const nextAllocations = { ...currentAllocations };

      submittedGroupKeys.forEach((groupKey) => {
        const currentGroupAllocations = nextAllocations[groupKey] || {};
        const nextGroupAllocations = { ...currentGroupAllocations };

        Object.entries(currentGroupAllocations).forEach(([siteKey, entry]) => {
          if (entry?.locked) return;

          nextGroupAllocations[siteKey] = {
            ...entry,
            locked: true,
            billingAllocationLockedAt:
              entry?.billingAllocationLockedAt || lockedAt,
          };
          hasChanges = true;
        });

        nextAllocations[groupKey] = nextGroupAllocations;
      });

      return hasChanges ? nextAllocations : currentAllocations;
    });
  }

  useEffect(() => {
    if (!activeCommonComplexBillingRow) {
      setAllocationDraftFee("");
      return;
    }

    setAllocationDraftFee(activeCommonComplexBillingRow.allocatedFee || "");
  }, [activeCommonComplexBillingRow]);

  function handleOpenCommonComplexAllocationModal(siteKey) {
    setAllocationSaveMessage("");
    setActiveAllocationSiteKey(siteKey);
  }

  function handleCloseCommonComplexAllocationModal() {
    setActiveAllocationSiteKey("");
    setAllocationDraftFee("");
    setAllocationSaveMessage("");
  }

  function handleCommonComplexAllocationInlineChange(siteKey, value) {
    if (!activeCommonComplexBillingGroupKey || !siteKey) return;
    const targetRow = commonComplexBillingRows.find(
      (row) => row.siteKey === siteKey,
    );
    if (
      activeBillingSection !== billingSections.REALLOCATION &&
      targetRow?.firstBillingCompleted
    ) {
      setAllocationSaveMessage(
        `${targetRow.screenCode} already has completed billing. Use Allocation/Reallocation for changes to its allocated fee.`,
      );
      return;
    }
    const normalizedFee = sanitizeAmount(value);

    setCommonComplexAllocations((currentAllocations) => {
      const currentGroup = currentAllocations[activeCommonComplexBillingGroupKey] || {};
      const currentEntry = currentGroup[siteKey] || {};
      if (
        currentEntry.locked &&
        activeBillingSection !== billingSections.REALLOCATION
      ) return currentAllocations;

      return {
        ...currentAllocations,
        [activeCommonComplexBillingGroupKey]: {
          ...currentGroup,
          [siteKey]: { ...currentEntry, allocatedFee: normalizedFee },
        },
      };
    });
  }

  async function persistCommonComplexBillingAllocation(
    groupKey,
    groupAllocations,
    rows,
    commonFee,
    savedAt,
  ) {
    if (!activeRecord?.backendSiteId) {
      throw new Error("Backend Site ID is missing for this billing record.");
    }

    const allocations = {};
    rows.forEach((row) => {
      const entry = groupAllocations[row.siteKey] || {};
      allocations[row.siteKey] = {
        ...entry,
        screenCode: row.screenCode,
        allocatedFee: normalizeValue(entry.allocatedFee || row.allocatedFee),
      };
    });

    const snapshot = {
      version: 1,
      groupKey,
      billingCode: normalizeValue(activeRecord.billingCode),
      complexCode: normalizeValue(activeRecord.complexCode),
      commonSubscriptionFee: formatAmountValue(commonFee),
      allocations,
      savedAt,
    };
    const updatedSite = await billingApiRequest(
      `/sites/${activeRecord.backendSiteId}`,
      {
        method: "PATCH",
        body: JSON.stringify({
          stage3Data: {
            ...activeRecord,
            commonComplexBillingAllocation: snapshot,
          },
        }),
      },
    );

    setStage3Records((currentRecords) =>
      currentRecords.map((record) =>
        record.backendSiteId === activeRecord.backendSiteId
          ? {
              ...record,
              ...(updatedSite?.stage3Data || {}),
              commonComplexBillingAllocation: snapshot,
            }
          : record,
      ),
    );

    return snapshot;
  }

  async function handleSaveAllCommonComplexAllocations() {
    if (!activeRecord || !activeCommonComplexBillingGroupKey) return;

    const rows = getCommonComplexBillingRows(
      activeRecord,
      stage3Records,
      commonComplexAllocations,
      commonScopeMembershipRecords,
    );
    const commonFee = parseAmountValue(
      resolveCommonComplexSubscriptionFee(
        activeRecord,
        stage3Records,
        commonComplexAllocations,
      ),
    );

    if (rows.length === 0 || commonFee === null) {
      alert("Common complex allocation could not be prepared.");
      return;
    }

    const completedRows = rows.filter((row) => row.firstBillingCompleted);
    if (completedRows.length > 0) {
      alert(
        `${completedRows.map((row) => row.screenCode).join(", ")} already has completed billing. Use Allocation/Reallocation for changes to its allocated fee.`,
      );
      return;
    }

    let total = 0;
    for (const row of rows) {
      const fee = parseAmountValue(row.allocatedFee);
      if (fee === null) {
        alert(
          `Allocation is required for ${row.screenCode}. All included screens must be allocated before saving.`,
        );
        return;
      }
      total += fee;
    }

    if (Math.abs(total - commonFee) > 0.01) {
      alert(
        `Site allocation total ${formatAmountValue(total)} does not match the Common Subscription Fee ${formatAmountValue(commonFee)}.`,
      );
      return;
    }

    const now = getCurrentLocalTimestamp();
    const currentGroup =
      commonComplexAllocations[activeCommonComplexBillingGroupKey] || {};
    const nextGroup = { ...currentGroup };

    rows.forEach((row) => {
      const currentEntry = nextGroup[row.siteKey] || {};
      if (currentEntry.locked) return;
      nextGroup[row.siteKey] = {
        ...currentEntry,
        allocatedFee: row.allocatedFee,
        billingAllocationCreatedAt:
          currentEntry.billingAllocationCreatedAt || now,
        billingAllocationUpdatedAt: now,
      };
    });

    try {
      await persistCommonComplexBillingAllocation(
        activeCommonComplexBillingGroupKey,
        nextGroup,
        rows,
        commonFee,
        now,
      );
    } catch (error) {
      alert(`Unable to save Complex billing allocation. ${error.message}`);
      return;
    }

    setCommonComplexAllocations((currentAllocations) => ({
      ...currentAllocations,
      [activeCommonComplexBillingGroupKey]: nextGroup,
    }));
    const affectedSiteKeys = new Set(rows.map((row) => row.siteKey));
    setStage3Records((currentRecords) =>
      currentRecords.map((record) =>
        affectedSiteKeys.has(getCommonComplexBillingSiteKey(record.screenCode))
          ? invalidateFirstTimeBillingValidation(record, "allocationRows")
          : record,
      ),
    );

    alert("Complex billing allocation saved successfully.");
  }

  async function handleSaveCommonComplexAllocation() {
    if (!activeCommonComplexBillingRow || !activeCommonComplexBillingGroupKey) {
      return;
    }

    if (activeCommonComplexBillingRow.locked) {
      handleCloseCommonComplexAllocationModal();
      return;
    }

    if (
      activeBillingSection !== billingSections.REALLOCATION &&
      activeCommonComplexBillingRow.firstBillingCompleted
    ) {
      setAllocationSaveMessage(
        `${activeCommonComplexBillingRow.screenCode} already has completed billing. Use Allocation/Reallocation for changes to its allocated fee.`,
      );
      return;
    }

    const normalizedFee = sanitizeAmount(allocationDraftFee);
    const parsedFee = parseAmountValue(normalizedFee);

    if (parsedFee === null) {
      setAllocationSaveMessage("Enter a valid Site-wise Billing Allocation.");
      return;
    }

    const now = getCurrentLocalTimestamp();
    const persistedFee = formatAmountValue(parsedFee);

    const currentGroupAllocations =
      commonComplexAllocations[activeCommonComplexBillingGroupKey] || {};
    const currentEntry = currentGroupAllocations[activeAllocationSiteKey] || {};
    const nextGroupAllocations = {
      ...currentGroupAllocations,
      [activeAllocationSiteKey]: {
        ...currentEntry,
        allocatedFee: persistedFee,
        billingAllocationCreatedAt:
          currentEntry.billingAllocationCreatedAt || now,
        billingAllocationUpdatedAt: now,
      },
    };
    const nextAllocations = {
      ...commonComplexAllocations,
      [activeCommonComplexBillingGroupKey]: nextGroupAllocations,
    };

    try {
      await persistCommonComplexBillingAllocation(
        activeCommonComplexBillingGroupKey,
        nextGroupAllocations,
        commonComplexBillingRows,
        parseAmountValue(
          resolveCommonComplexSubscriptionFee(
            activeRecord,
            stage3Records,
            nextAllocations,
          ),
        ) || 0,
        now,
      );
    } catch (error) {
      setAllocationSaveMessage(`Unable to save allocation. ${error.message}`);
      return;
    }

    setCommonComplexAllocations(nextAllocations);
    setStage3Records((currentRecords) =>
      currentRecords.map((record) =>
        getCommonComplexBillingSiteKey(record.screenCode) ===
        activeAllocationSiteKey
          ? invalidateFirstTimeBillingValidation(record, "allocationRows")
          : record,
      ),
    );

    setAllocationSaveMessage("Allocation saved.");
    handleCloseCommonComplexAllocationModal();
  }

  function handleDownloadBillingRecordsExcel() {
    if (filteredSubmittedBillingRecords.length === 0) {
      alert("No Billing Records are available for download.");
      return;
    }

    const exportRows = filteredSubmittedBillingRecords.map((record) => {
      const commonPricing = isCommonComplexBillingRecord(record);

      return {
        "Customer Code/Billing Code": normalizeValue(record.billingCode),
        "Complex Code": normalizeValue(record.complexCode),
        "Screen Code / Scope": normalizeValue(
          record.screenCode || record.siteScope,
        ),
        "Pricing Mode": commonPricing
          ? "Common"
          : normalizeValue(record.pricingMethod),
        "Effective Screen Fee": getBillingRecordEffectiveScreenFee(
          record,
          commonComplexAllocations,
        ),
        "Billing Period From": normalizeValue(record.billingPeriodFrom),
        "Billing Period To": normalizeValue(record.billingPeriodTo),
        "Billing Amount Before GST": normalizeValue(
          record.billingAmountBeforeGST,
        ),
        "Submission Status": normalizeValue(record.submissionStatus),
        "Invoice Number": normalizeValue(record.invoiceNumber),
        "Invoice Date": normalizeValue(record.invoiceDate),
        "Invoice Amount Before GST": normalizeValue(
          record.invoiceAmountBeforeGST,
        ),
        "Invoice Status": normalizeValue(record.invoiceStatus),
        "Payment Status": normalizeValue(record.paymentStatus),
        "Payment Received Date": normalizeValue(record.paymentReceivedDate),
        "Payment Remarks": normalizeValue(record.paymentRemarks),
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Billing Records");

    const commonPricingExportRows = billingRecordsCommonPricingSummaries.flatMap(
      (summary) =>
        summary.screenRows.map((row) => ({
          "Billing ID / Complex":
            summary.complexCode || summary.billingCode || summary.groupKey,
          "Screen Code": row.screenCode,
          "Screen Fee": row.allocatedFee,
          "Screen Status": row.status,
          "Included in Current Invoice": row.isBillable ? "Yes" : "No",
          "Current Combined Invoice Amount":
            summary.currentCombinedBillingAmount,
          "Configured Common Total": summary.configuredCommonTotal,
        })),
    );

    if (commonPricingExportRows.length > 0) {
      const commonWorksheet = XLSX.utils.json_to_sheet(
        commonPricingExportRows,
      );
      XLSX.utils.book_append_sheet(
        workbook,
        commonWorksheet,
        "Common Pricing Summary",
      );
    }

    XLSX.writeFile(
      workbook,
      `Billing_Records_${new Date().toISOString().slice(0, 10)}.xlsx`,
    );
  }

  function getStage3RecordForBillingRecord(record) {
    if (!record) {
      return null;
    }

    const sourceRecordId = normalizeValue(record.sourceRecordId);
    const screenCode = normalizeValue(record.screenCode || record.siteScope);

    return (
      stage3Records.find(
        (site) =>
          sourceRecordId &&
          normalizeValue(site.recordId) === sourceRecordId,
      ) ||
      stage3Records.find(
        (site) =>
          screenCode &&
          normalizeValue(site.screenCode) === screenCode,
      ) ||
      null
    );
  }

  function persistBillingRecordInactiveStatus(forcedNextStatus = "") {
    if (!activeBillingRecord || !billingRecordDraft) {
      return false;
    }

    const nextStatus =
      normalizeValue(forcedNextStatus) ||
      normalizeValue(billingRecordDraft.billingStatus) ||
      "Active";

    if (nextStatus !== "Inactive") {
      return false;
    }

    const now = new Date().toISOString();
    const inactiveEffectiveDate = formatLocalDateValue(new Date());
    const sourceRecordId = normalizeValue(activeBillingRecord.sourceRecordId);
    const screenCode = normalizeValue(
      activeBillingRecord.screenCode || activeBillingRecord.siteScope,
    );
    const billingRecordId = normalizeValue(
      activeBillingRecord.billingRecordId,
    );

    let stage3WasUpdated = false;
    let billingRecordWasUpdated = false;

    setStage3Records((currentRecords) => {
      let matched = false;

      const updatedRecords = currentRecords.map((record) => {
        const sameSource =
          sourceRecordId && normalizeValue(record.recordId) === sourceRecordId;
        const sameScreen =
          screenCode && normalizeValue(record.screenCode) === screenCode;
        const sameBillingIdentity =
          normalizeValue(record.billingCode) ===
            normalizeValue(activeBillingRecord.billingCode) &&
          normalizeValue(record.screenName) ===
            normalizeValue(activeBillingRecord.screenName);

        if (!sameSource && !sameScreen && !sameBillingIdentity) {
          return record;
        }

        matched = true;
        stage3WasUpdated = true;

        const alreadyInactive =
          normalizeValue(record.billingLifecycleStatus) === "Inactive";

        return {
          ...record,
          billingLifecycleStatus: "Inactive",
          inactiveSince: record.inactiveSince || now,
          inactiveEffectiveDate:
            record.inactiveEffectiveDate || inactiveEffectiveDate,
          billingLifecycleReason:
            normalizeValue(record.billingLifecycleReason) || "",
          closureEffectiveDate: record.closureEffectiveDate || "",
          closureRemarks: record.closureRemarks || "",
          closureUpdatedAt: now,
          closureHistory: alreadyInactive
            ? Array.isArray(record.closureHistory)
              ? record.closureHistory
              : []
            : [
                ...(Array.isArray(record.closureHistory)
                  ? record.closureHistory
                  : []),
                {
                  eventId: `inactive-${Date.now()}-${Math.random()
                    .toString(36)
                    .slice(2, 7)}`,
                  previousStatus: "Active",
                  newStatus: "Inactive",
                  effectiveDate: inactiveEffectiveDate,
                  remarks:
                    "Billing Records status changed to Inactive. Pending billing, if any, must complete before movement to Paused / Inactive.",
                  updatedAt: now,
                },
              ],
        };
      });

      if (matched) {
        return updatedRecords;
      }

      stage3WasUpdated = true;

      return [
        ...updatedRecords,
        {
          ...activeBillingRecord,
          recordId:
            sourceRecordId ||
            `lifecycle-${screenCode || Date.now()}-${Math.random()
              .toString(36)
              .slice(2, 7)}`,
          sourceRecordId:
            sourceRecordId ||
            normalizeValue(activeBillingRecord.sourceRecordId),
          billingLifecycleStatus: "Inactive",
          billingLifecycleReason: "",
          inactiveSince: now,
          inactiveEffectiveDate,
          closureEffectiveDate: "",
          closureRemarks: "",
          closureUpdatedAt: now,
          closureHistory: [
            {
              eventId: `inactive-${Date.now()}-${Math.random()
                .toString(36)
                .slice(2, 7)}`,
              previousStatus: "Active",
              newStatus: "Inactive",
              effectiveDate: inactiveEffectiveDate,
              remarks:
                "Billing Records status changed to Inactive. Pending billing, if any, must complete before movement to Paused / Inactive.",
              updatedAt: now,
            },
          ],
        },
      ];
    });

    setBillingRecords((currentRecords) =>
      currentRecords.map((record) => {
        const sameBillingRecord =
          billingRecordId &&
          normalizeValue(record.billingRecordId) === billingRecordId;
        const sameSource =
          sourceRecordId &&
          normalizeValue(record.sourceRecordId) === sourceRecordId;
        const sameScreen =
          screenCode &&
          normalizeValue(record.screenCode || record.siteScope) === screenCode;
        const sameVisibleIdentity =
          normalizeValue(record.billingCode) ===
            normalizeValue(activeBillingRecord.billingCode) &&
          normalizeValue(record.screenName) ===
            normalizeValue(activeBillingRecord.screenName);

        if (
          !sameBillingRecord &&
          !sameSource &&
          !sameScreen &&
          !sameVisibleIdentity
        ) {
          return record;
        }

        billingRecordWasUpdated = true;

        const alreadyInactive =
          normalizeValue(record.billingLifecycleStatus) === "Inactive";

        return {
          ...record,
          billingLifecycleStatus: "Inactive",
          billingStatusReason: normalizeValue(record.billingStatusReason),
          inactiveSince: record.inactiveSince || now,
          inactiveEffectiveDate:
            record.inactiveEffectiveDate || inactiveEffectiveDate,
          lastUpdatedAt: now,
          auditTrail: alreadyInactive
            ? Array.isArray(record.auditTrail)
              ? record.auditTrail
              : []
            : [
                ...(Array.isArray(record.auditTrail)
                  ? record.auditTrail
                  : []),
                createBillingAuditEntry({
                  event: "Billing Status Changed",
                  field: "billingStatus",
                  previousValue: "Active",
                  newValue: "Inactive",
                  remarks:
                    "Status changed to Inactive. Pending billing, if any, will complete before movement to Paused / Inactive.",
                  timestamp: now,
                }),
              ],
        };
      }),
    );

    setBillingRecordDraft((currentDraft) =>
      currentDraft
        ? {
            ...currentDraft,
            billingStatus: "Inactive",
          }
        : currentDraft,
    );

    return true;
  }

  function handleBillingStatusDraftChange(nextStatus) {
    if (!billingRecordEditMode || !billingRecordDraft) {
      return;
    }

    if (nextStatus === "Inactive") {
      const confirmed = window.confirm(
        "Billing Status will change to Inactive now. Existing Billing Records will remain unchanged. If billing is still pending, the site will move to Paused / Inactive only after the pending billing is completed. Continue?",
      );

      if (!confirmed) {
        return;
      }

      setBillingRecordDraft((currentDraft) =>
        currentDraft
          ? {
              ...currentDraft,
              billingStatus: "Inactive",
            }
          : currentDraft,
      );

      const statusUpdated =
        persistBillingRecordInactiveStatus("Inactive");

      setBillingRecordSaveMessage(
        statusUpdated
          ? "Billing Status updated to Inactive. Any pending billing must complete before the site moves to Paused / Inactive."
          : "Billing Status is already Inactive.",
      );
      return;
    }

    setBillingRecordDraft((currentDraft) =>
      currentDraft
        ? {
            ...currentDraft,
            billingStatus: "Active",
          }
        : currentDraft,
    );
    setBillingRecordSaveMessage("");
  }

  function handleOpenBillingRecord(recordId) {
    setBillingRecordSaveMessage("");
    setBillingRecordEditMode(false);
    setActiveBillingRecordId(recordId);
    setActiveWorkspace(workspaceModes.BILLING);
  }

  function handleEditBillingRecord(recordId) {
    const record = submittedBillingRecords.find(
      (candidate) => candidate.billingRecordId === recordId,
    );
    const currentSite = resolveBillingRecordSite(
      record,
      billingRecordSiteCandidates,
    );
    if (!record || !currentSite) {
      alert("Unable to resolve the current Site for editing this Billing Record.");
      return;
    }

    setBillingRecordSaveMessage("");
    setBillingRecordEditMode(true);
    setActiveBillingRecordId(recordId);
    setActiveWorkspace(workspaceModes.BILLING);
  }

  function handleCloseBillingRecord() {
    setActiveBillingRecordId("");
    setBillingRecordDraft(null);
    setBillingRecordSaveMessage("");
    setBillingRecordEditMode(false);
    setBillingRecordFinancialYear("");
    setBillingRecordMonthFilter("All");
    setBillingRecordSelectedEntryKey("");
    setBillingRecordMonthDetailMode("view");
    setBillingRecordManualPaymentDraft(null);
    setBillingRecordManualPaymentMessage("");
  }

  function handleOpenBillingHistory(record) {
    const screenCode = normalizeValue(record?.screenCode || record?.siteScope);

    if (!screenCode) {
      alert("Screen Code is not available for this Billing Record.");
      return;
    }

    setBillingHistoryScreenCode(screenCode);
    setActiveBillingRecordId("");
    setBillingRecordDraft(null);
    setBillingRecordSaveMessage("");
  }

  function handleCloseBillingHistory() {
    setBillingHistoryScreenCode("");
  }

  function updateBillingRecordDraft(field, value) {
    setBillingRecordDraft((currentDraft) =>
      currentDraft
        ? {
            ...currentDraft,
            [field]: value,
          }
        : currentDraft,
    );
  }

  function updateBillingInvoiceEntry(entryId, field, value) {
    setBillingRecordDraft((currentDraft) => {
      if (!currentDraft) {
        return currentDraft;
      }

      return {
        ...currentDraft,
        invoiceEntries: (currentDraft.invoiceEntries || []).map((entry) => {
          if (entry.entryId !== entryId) {
            return entry;
          }

          const nextEntry = {
            ...entry,
            [field]: value,
          };

          if (
            field === "invoiceNumber" ||
            field === "invoiceDate" ||
            field === "invoiceAmountBeforeGST"
          ) {
            nextEntry.invoiceStatus = getBillingRecordInvoiceStatus(nextEntry);
          }

          return nextEntry;
        }),
      };
    });
  }

  function handleBillingRecordValidatedDateBlur(field) {
    return (event) => {
      const nextValue = event.target.value;

      if (nextValue && !isValidDateValue(nextValue)) {
        alert("Please enter a valid calendar date.");
      }
    };
  }


  function openBillingRecordMonthDetails(row, mode) {
    const requestedMode =
      normalizeValue(row?.submittedPeriodTreatment).toLowerCase() === "waived"
        ? "view"
        : mode;

    setBillingRecordSelectedEntryKey(row.rowKey);
    setBillingRecordMonthDetailMode(requestedMode);
    setIsErpMonthlyFullscreen(true);
    setBillingRecordManualPaymentMessage("");

    if (requestedMode === "edit") {
      setBillingRecordManualPaymentDraft({
        invoiceNumber: normalizeValue(row.invoiceNumber),
        invoiceDate: normalizeValue(row.invoiceDate),
        invoiceAmountBeforeGST: normalizeValue(
          row.invoiceAmountBeforeGST || row.billingAmount,
        ),
        receiptNumber: normalizeValue(row.receiptNumber),
        receiptDate: normalizeValue(row.receiptDate),
        receivedAmount: normalizeValue(row.receivedAmount),
        paymentStatus: row.lifecycleBillable === false
          ? "Paused / Not Billable"
          : normalizeValue(row.paymentStatus) ||
            (normalizeValue(row.invoiceNumber) ? "Pending" : "Not Invoiced"),
        paymentRemarks: row.lifecycleBillable === false
          ? normalizeValue(row.lifecycleRemarks) || normalizeValue(row.paymentRemarks)
          : normalizeValue(row.paymentRemarks),
      });
    } else {
      setBillingRecordManualPaymentDraft(null);
    }
  }

  function updateBillingRecordManualPaymentDraft(field, value) {
    setBillingRecordManualPaymentDraft((currentDraft) =>
      currentDraft
        ? {
            ...currentDraft,
            [field]: value,
          }
        : currentDraft,
    );
    setBillingRecordManualPaymentMessage("");
  }

  async function handleSaveManualBillingRecordPayment() {
    if (
      !activeBillingRecord ||
      !selectedBillingRecordFinancialRow ||
      !billingRecordManualPaymentDraft
    ) {
      return;
    }

    const selectedRow = selectedBillingRecordFinancialRow;
    if (isSelectedBillingMonthWaived) {
      setBillingRecordManualPaymentMessage(
        "Waived billing periods are read-only.",
      );
      return;
    }
    const invoiceNumber = normalizeValue(
      billingRecordManualPaymentDraft.invoiceNumber,
    );
    const invoiceDate = normalizeValue(
      billingRecordManualPaymentDraft.invoiceDate,
    );
    const invoiceAmountBeforeGST = normalizeValue(
      billingRecordManualPaymentDraft.invoiceAmountBeforeGST,
    );

    if (!invoiceNumber) {
      setBillingRecordManualPaymentMessage("Invoice Number is required.");
      return;
    }

    const invoiceConflict = findInvoiceConflict({
      candidateRecord: activeBillingRecord,
      candidateEntry: {
        ...selectedRow,
        invoiceNumber,
      },
      billingRecords,
      ignoredBillingRecordId: normalizeValue(
        activeBillingRecord.billingRecordId,
      ),
    });

    if (invoiceConflict) {
      setBillingRecordManualPaymentMessage(invoiceConflict);
      return;
    }

    if (!invoiceDate) {
      setBillingRecordManualPaymentMessage("Invoice Date is required.");
      return;
    }

    if (!isValidDateValue(invoiceDate)) {
      setBillingRecordManualPaymentMessage("Please enter a valid Invoice Date.");
      return;
    }

    if (
      !invoiceAmountBeforeGST ||
      parseAmountValue(invoiceAmountBeforeGST) === null
    ) {
      setBillingRecordManualPaymentMessage(
        "Invoice Amount Before GST is required and must be a valid amount.",
      );
      return;
    }

    const receiptNumber = normalizeValue(
      billingRecordManualPaymentDraft.receiptNumber,
    );
    const receiptDate = normalizeValue(
      billingRecordManualPaymentDraft.receiptDate,
    );
    const receivedAmount = normalizeValue(
      billingRecordManualPaymentDraft.receivedAmount,
    );
    const paymentStatus =
      normalizeValue(billingRecordManualPaymentDraft.paymentStatus) || "Pending";
    const paymentRemarks = normalizeValue(
      billingRecordManualPaymentDraft.paymentRemarks,
    );

    if (receiptDate && !isValidDateValue(receiptDate)) {
      setBillingRecordManualPaymentMessage("Please enter a valid Receipt Date.");
      return;
    }

    if (
      isPaymentReceivedStatus(paymentStatus) &&
      receiptDate &&
      isFutureLocalDateValue(receiptDate)
    ) {
      setBillingRecordManualPaymentMessage(
        "Receipt Date must not be a future date. Future receipt date - payment not yet effective",
      );
      return;
    }

    if (
      receivedAmount &&
      (!Number.isFinite(Number(receivedAmount)) || Number(receivedAmount) < 0)
    ) {
      setBillingRecordManualPaymentMessage(
        "Received Amount must be a valid non-negative amount.",
      );
      return;
    }

    if (
      ["Paid", "Partially Paid"].includes(paymentStatus) &&
      !receivedAmount
    ) {
      setBillingRecordManualPaymentMessage(
        "Received Amount is required when Payment Status is Paid or Partially Paid.",
      );
      return;
    }

    if (isPaymentReceivedStatus(paymentStatus) && !receiptDate) {
      setBillingRecordManualPaymentMessage(
        "Receipt Date is required when Payment Status is Paid or Payment Received.",
      );
      return;
    }

    const targetBillingRecordId =
      normalizeValue(selectedRow.sourceBillingRecordId) ||
      normalizeValue(activeBillingRecord.billingRecordId);

    const now = new Date().toISOString();

    const persistedSite = stage3Records.find(
      (record) =>
        normalizeValue(record?.backendSiteId) ===
        normalizeValue(activeBillingRecord?.backendSiteId),
    );
    const persistedPeriods = buildBillingInvoicePeriods(activeBillingRecord);
    const persistedTargetPeriod =
      persistedPeriods.find(
        (entry) => entry.entryId === normalizeValue(selectedRow.entryId),
      ) ||
      persistedPeriods.find(
        (entry) =>
          normalizeValue(entry.periodFrom) ===
          normalizeValue(selectedRow.periodFrom),
      );
    const persistedFirstTimeRecord = persistedTargetPeriod
      ? {
          ...activeBillingRecord,
          invoiceEntries: persistedPeriods.map((entry) =>
            entry.entryId === persistedTargetPeriod.entryId
              ? {
                  ...entry,
                  invoiceNumber,
                  invoiceDate,
                  invoiceAmountBeforeGST,
                  receiptNumber,
                  receiptDate,
                  receivedAmount,
                  paymentStatus,
                  paymentRemarks,
                  manualPaymentSource: "Manual Entry",
                  manualPaymentUpdatedAt: now,
                }
              : entry,
          ),
          lastUpdatedAt: now,
        }
      : null;

    const isFirstTimeBillingRecord =
      normalizeValue(activeBillingRecord?.billingCycleType).toLowerCase() ===
      "first time";

    if (
      isFirstTimeBillingRecord &&
      (!persistedSite?.backendSiteId || !persistedFirstTimeRecord)
    ) {
      setBillingRecordManualPaymentMessage(
        "Unable to persist invoice and payment details because the canonical Site record is unavailable.",
      );
      return;
    }

    if (isFirstTimeBillingRecord) {
      try {
        await billingApiRequest(`/sites/${persistedSite.backendSiteId}`, {
          method: "PATCH",
          body: JSON.stringify({
            stage3Data: {
              ...persistedSite,
              firstTimeBillingRecord: persistedFirstTimeRecord,
            },
          }),
        });
      } catch (error) {
        setBillingRecordManualPaymentMessage(
          error instanceof Error
            ? error.message
            : "Unable to persist invoice and payment details.",
        );
        return;
      }
    }

    let saved = false;

    setBillingRecords((currentRecords) =>
      currentRecords.map((record) => {
        if (record.billingRecordId !== targetBillingRecordId) {
          return record;
        }

        const periods = buildBillingInvoicePeriods(record);
        const targetEntryId = normalizeValue(selectedRow.entryId);
        const targetPeriod =
          periods.find((entry) => entry.entryId === targetEntryId) ||
          periods.find(
            (entry) =>
              normalizeValue(entry.periodFrom) ===
              normalizeValue(selectedRow.periodFrom),
          );

        if (!targetPeriod) {
          return record;
        }

        const previous = {
          invoiceNumber: normalizeValue(targetPeriod.invoiceNumber),
          invoiceDate: normalizeValue(targetPeriod.invoiceDate),
          invoiceAmountBeforeGST: normalizeValue(
            targetPeriod.invoiceAmountBeforeGST,
          ),
          receiptNumber: normalizeValue(targetPeriod.receiptNumber),
          receiptDate: normalizeValue(targetPeriod.receiptDate),
          receivedAmount: normalizeValue(targetPeriod.receivedAmount),
          paymentStatus:
            normalizeValue(targetPeriod.paymentStatus) || "Pending",
          paymentRemarks: normalizeValue(targetPeriod.paymentRemarks),
        };

        const nextEntry = {
          ...targetPeriod,
          invoiceNumber,
          invoiceDate,
          invoiceAmountBeforeGST,
          receiptNumber,
          receiptDate,
          receivedAmount,
          paymentStatus,
          paymentRemarks,
          manualPaymentSource: "Manual Entry",
          manualPaymentUpdatedAt: now,
        };

        const nextEntries = periods.map((entry) =>
          entry.entryId === targetPeriod.entryId ? nextEntry : entry,
        );

        const changes = [
          ["invoiceNumber", previous.invoiceNumber, invoiceNumber],
          ["invoiceDate", previous.invoiceDate, invoiceDate],
          [
            "invoiceAmountBeforeGST",
            previous.invoiceAmountBeforeGST,
            invoiceAmountBeforeGST,
          ],
          ["receiptNumber", previous.receiptNumber, receiptNumber],
          ["receiptDate", previous.receiptDate, receiptDate],
          ["receivedAmount", previous.receivedAmount, receivedAmount],
          ["paymentStatus", previous.paymentStatus, paymentStatus],
          ["paymentRemarks", previous.paymentRemarks, paymentRemarks],
        ].filter(
          ([, previousValue, nextValue]) =>
            normalizeValue(previousValue) !== normalizeValue(nextValue),
        );

        if (changes.length === 0) {
          saved = true;
          return record;
        }

        const auditEntries = changes.map(
          ([field, previousValue, nextValue]) =>
            createBillingAuditEntry({
              event: `${selectedRow.billingMonth || "Monthly"} Manual Payment Entry`,
              field,
              previousValue,
              newValue: nextValue,
              remarks:
                paymentRemarks ||
                "Manual receipt/payment fallback used because ERP sync data was unavailable.",
              timestamp: now,
            }),
        );

        saved = true;

        return {
          ...record,
          invoiceEntries: nextEntries,
          invoiceNumber:
            record.billingRecordId === activeBillingRecord.billingRecordId
              ? invoiceNumber
              : record.invoiceNumber,
          invoiceDate:
            record.billingRecordId === activeBillingRecord.billingRecordId
              ? invoiceDate
              : record.invoiceDate,
          invoiceAmountBeforeGST:
            record.billingRecordId === activeBillingRecord.billingRecordId
              ? invoiceAmountBeforeGST
              : record.invoiceAmountBeforeGST,
          receiptNumber:
            record.billingRecordId === activeBillingRecord.billingRecordId
              ? receiptNumber
              : record.receiptNumber,
          receiptDate:
            record.billingRecordId === activeBillingRecord.billingRecordId
              ? receiptDate
              : record.receiptDate,
          receivedAmount:
            record.billingRecordId === activeBillingRecord.billingRecordId
              ? receivedAmount
              : record.receivedAmount,
          paymentStatus:
            record.billingRecordId === activeBillingRecord.billingRecordId
              ? paymentStatus
              : record.paymentStatus,
          paymentRemarks:
            record.billingRecordId === activeBillingRecord.billingRecordId
              ? paymentRemarks
              : record.paymentRemarks,
          manualPaymentSource: "Manual Entry",
          manualPaymentUpdatedAt: now,
          lastUpdatedAt: now,
          auditTrail: [
            ...(Array.isArray(record.auditTrail) ? record.auditTrail : []),
            ...auditEntries,
          ],
        };
      }),
    );

    if (!saved) {
      setBillingRecordManualPaymentMessage(
        "Unable to locate the monthly billing record for this entry.",
      );
      return;
    }

    setBillingRecordManualPaymentMessage(
      "Manual receipt/payment details saved successfully.",
    );
    setBillingRecordMonthDetailMode("view");
    setBillingRecordManualPaymentDraft(null);
  }

  function handleSaveBillingRecord() {
    if (!activeBillingRecord || !billingRecordDraft) {
      return;
    }

    const nextInvoiceEntries = Array.isArray(billingRecordDraft.invoiceEntries)
      ? billingRecordDraft.invoiceEntries.map((entry) => ({
          ...entry,
          invoiceNumber: normalizeValue(entry.invoiceNumber),
          invoiceDate: normalizeValue(entry.invoiceDate),
          invoiceAmountBeforeGST: normalizeValue(entry.invoiceAmountBeforeGST),
          invoiceStatus: getBillingRecordInvoiceStatus(entry),
          paymentStatus: normalizeValue(entry.paymentStatus) || "Pending",
          paymentReceivedDate: normalizeValue(entry.paymentReceivedDate),
          paymentRemarks: normalizeValue(entry.paymentRemarks),
        }))
      : [];

    const seenInvoiceReferences = new Set();
    for (const entry of nextInvoiceEntries) {
      const invoiceNumber = normalizeInvoiceReference(entry.invoiceNumber);
      if (!invoiceNumber) continue;

      const invoiceKey = `${normalizeInvoiceReference(
        activeBillingRecord.billingCode,
      )}::${invoiceNumber}`;
      if (seenInvoiceReferences.has(invoiceKey)) {
        alert(
          `Invoice Number ${entry.invoiceNumber} has already been used for this Billing Code.`,
        );
        return;
      }
      seenInvoiceReferences.add(invoiceKey);

      const invoiceConflict = findInvoiceConflict({
        candidateRecord: activeBillingRecord,
        candidateEntry: entry,
        billingRecords,
        ignoredBillingRecordId: normalizeValue(
          activeBillingRecord.billingRecordId,
        ),
      });
      if (invoiceConflict) {
        alert(invoiceConflict);
        return;
      }
    }

    for (const entry of nextInvoiceEntries) {
      if (entry.invoiceDate && !isValidDateValue(entry.invoiceDate)) {
        alert(`${entry.billingMonth}: Please enter a valid Invoice Date.`);
        return;
      }

      if (
        entry.paymentReceivedDate &&
        !isValidDateValue(entry.paymentReceivedDate)
      ) {
        alert(
          `${entry.billingMonth}: Please enter a valid Payment Received Date.`,
        );
        return;
      }

      if (
        isPaymentReceivedStatus(entry.paymentStatus) &&
        entry.paymentReceivedDate &&
        isFutureLocalDateValue(entry.paymentReceivedDate)
      ) {
        alert(
          `${entry.billingMonth}: Payment Received Date must not be a future date. Future receipt date - payment not yet effective`,
        );
        return;
      }

      if (
        isPaymentReceivedStatus(entry.paymentStatus) &&
        !entry.paymentReceivedDate
      ) {
        alert(
          `${entry.billingMonth}: Payment Received Date is required when status is Paid or Payment Received.`,
        );
        return;
      }
    }

    const now = new Date().toISOString();
    const previousEntries = buildBillingInvoicePeriods(activeBillingRecord);
    const previousById = new Map(
      previousEntries.map((entry) => [entry.entryId, entry]),
    );
    const auditEntries = [];

    nextInvoiceEntries.forEach((entry) => {
      const previousEntry = previousById.get(entry.entryId) || {};
      const monthLabel = entry.billingMonth || entry.entryId;

      [
        ["invoiceNumber", previousEntry.invoiceNumber, entry.invoiceNumber],
        ["invoiceDate", previousEntry.invoiceDate, entry.invoiceDate],
        [
          "invoiceAmountBeforeGST",
          previousEntry.invoiceAmountBeforeGST,
          entry.invoiceAmountBeforeGST,
        ],
        ["invoiceStatus", previousEntry.invoiceStatus, entry.invoiceStatus],
        ["paymentStatus", previousEntry.paymentStatus, entry.paymentStatus],
        [
          "paymentReceivedDate",
          previousEntry.paymentReceivedDate,
          entry.paymentReceivedDate,
        ],
        ["paymentRemarks", previousEntry.paymentRemarks, entry.paymentRemarks],
      ]
        .filter(
          ([, previousValue, nextValue]) =>
            normalizeValue(previousValue) !== normalizeValue(nextValue),
        )
        .forEach(([field, previousValue, nextValue]) => {
          auditEntries.push(
            createBillingAuditEntry({
              event: `${monthLabel} Billing Record Updated`,
              field,
              previousValue,
              newValue: nextValue,
              remarks:
                entry.paymentRemarks ||
                `${monthLabel} invoice/payment details updated.`,
              timestamp: now,
            }),
          );
        });
    });

    const invoiceCount = nextInvoiceEntries.filter(
      (entry) => entry.invoiceStatus === "Invoiced",
    ).length;
    const paidCount = nextInvoiceEntries.filter(
      (entry) => entry.paymentStatus === "Paid",
    ).length;
    const partialPaidCount = nextInvoiceEntries.filter(
      (entry) => entry.paymentStatus === "Partially Paid",
    ).length;

    const aggregateInvoiceStatus =
      nextInvoiceEntries.length > 0 && invoiceCount === nextInvoiceEntries.length
        ? "Invoiced"
        : invoiceCount > 0
          ? "Partially Invoiced"
          : "Pending Invoice";

    const aggregatePaymentStatus =
      nextInvoiceEntries.length > 0 && paidCount === nextInvoiceEntries.length
        ? "Paid"
        : paidCount > 0 || partialPaidCount > 0
          ? "Partially Paid"
          : "Pending";

    const firstEntry = nextInvoiceEntries[0] || {};

    setBillingRecords((currentRecords) =>
      currentRecords.map((record) =>
        record.billingRecordId === activeBillingRecord.billingRecordId
          ? {
              ...record,
              invoiceEntries: nextInvoiceEntries,
              invoiceNumber: firstEntry.invoiceNumber || "",
              invoiceDate: firstEntry.invoiceDate || "",
              invoiceAmountBeforeGST: firstEntry.invoiceAmountBeforeGST || "",
              invoiceStatus: aggregateInvoiceStatus,
              paymentStatus: aggregatePaymentStatus,
              paymentReceivedDate: firstEntry.paymentReceivedDate || "",
              paymentRemarks: firstEntry.paymentRemarks || "",
              submissionStatus:
                record.submissionStatus || "Submitted to Billing Team",
              submittedAt: record.submittedAt || now,
              createdAt: record.createdAt || now,
              lastUpdatedAt: now,
              auditTrail:
                auditEntries.length > 0
                  ? [
                      ...(Array.isArray(record.auditTrail)
                        ? record.auditTrail
                        : []),
                      ...auditEntries,
                    ]
                  : record.auditTrail,
            }
          : record,
      ),
    );

    setBillingRecordSaveMessage(
      auditEntries.length > 0
        ? "Billing record saved successfully. Monthly Invoice Details and Audit History updated."
        : "No billing record changes to save.",
    );

    setBillingRecordDraft((currentDraft) =>
      currentDraft
        ? {
            ...currentDraft,
            invoiceEntries: nextInvoiceEntries,
            invoiceNumber: firstEntry.invoiceNumber || "",
            invoiceDate: firstEntry.invoiceDate || "",
            invoiceAmountBeforeGST: firstEntry.invoiceAmountBeforeGST || "",
            invoiceStatus: aggregateInvoiceStatus,
            paymentStatus: aggregatePaymentStatus,
            paymentReceivedDate: firstEntry.paymentReceivedDate || "",
            paymentRemarks: firstEntry.paymentRemarks || "",
            submittedAt: currentDraft.submittedAt || now,
            createdAt: currentDraft.createdAt || now,
            lastUpdatedAt: now,
          }
        : currentDraft,
    );
  }

  function handleSaveBillingInternalEdit() {
    if (!activeBillingRecord || !billingRecordDraft) {
      return;
    }

    const movedToInactive = persistBillingRecordInactiveStatus();

    const previousRemarks = normalizeValue(activeBillingRecord.billingRemarks);
    const nextRemarks = normalizeValue(billingRecordDraft.billingRemarks);

    if (previousRemarks === nextRemarks) {
      setBillingRecordSaveMessage(
        movedToInactive
          ? "Billing Status updated to Inactive. Pending billing, if any, must complete before movement to Paused / Inactive."
          : "No Billing INC changes to save.",
      );
      return;
    }

    const now = new Date().toISOString();
    const auditEntry = createBillingAuditEntry({
      event: "Billing INC Record Edited",
      field: "billingRemarks",
      previousValue: previousRemarks,
      newValue: nextRemarks,
      remarks: "Internal Billing INC remark updated.",
      timestamp: now,
    });

    setBillingRecords((currentRecords) =>
      currentRecords.map((record) =>
        record.billingRecordId === activeBillingRecord.billingRecordId
          ? {
              ...record,
              billingRemarks: nextRemarks,
              lastUpdatedAt: now,
              auditTrail: [
                ...(Array.isArray(record.auditTrail) ? record.auditTrail : []),
                auditEntry,
              ],
            }
          : record,
      ),
    );

    setBillingRecordDraft((currentDraft) =>
      currentDraft
        ? {
            ...currentDraft,
            billingRemarks: nextRemarks,
            lastUpdatedAt: now,
          }
        : currentDraft,
    );

    setBillingRecordSaveMessage(
      "Billing INC internal remarks saved. ERP financial values remain read-only.",
    );
  }

  function handleOpenVerificationRecord(recordId) {
    setFirstTimeValidationMessage("");

    const selectedRecord =
      stage3Records.find((record) => record.recordId === recordId) ||
      verificationRows.find((record) => record.recordId === recordId);

    if (!selectedRecord) {
      return;
    }

    setActiveRecordId(selectedRecord.recordId);
    setPendingBillingNavigationTarget(
      isCommonComplexBillingRecord(selectedRecord) &&
      !isVariableSubscriptionRecord(selectedRecord)
        ? "common-allocation"
        : "billing-information",
    );
  }

  function handleOpenRecurringRecord(recordId) {
    const selectedRecord = recurringRows.find(
      (record) => record.recordId === recordId,
    );

    if (!selectedRecord) {
      return;
    }

    const savedHold =
      selectedRecord?.recurringBillingHolds?.[
        selectedRecord.recurringCycle?.monthKey
      ] || null;

    setActiveRecordId(selectedRecord.recordId);
    setRecurringDecision(savedHold?.status === "Hold" ? "Hold" : "Process");
    setRecurringRemarks(savedHold?.remarks || "");
    setPendingBillingNavigationTarget("recurring-workspace");
  }

  function handleRecurringDecisionSubmit() {
    if (!activeRecurringRecord) {
      return;
    }

    const recurringCycle = activeRecurringRecord.recurringCycle || {};

    if (
      recurringDecision === "Process" &&
      recurringCycle.priceChangeBlocked === true
    ) {
      const blockedLabel =
        recurringCycle.priceChangeBlockLabel ||
        activeRecurringRecord.screenName ||
        activeRecurringRecord.screenCode ||
        activeRecurringRecord.complexCode ||
        "This site/complex";
      const proceed = window.confirm(
        `${blockedLabel} is currently under Price Change & Reallocation. Recurring Billing cannot be processed using the previous commercial value.

To proceed with the remaining eligible sites/complexes while excluding this site/complex, select Yes/OK. Select No/Cancel to go back and validate the site/complex in Price Change & Reallocation.`,
      );

      if (proceed) {
        setActiveRecordId("");
        setRecurringDecision("Process");
        setRecurringRemarks("");
      }
      return;
    }
    const isCombinedCommonBilling =
      activeRecurringRecord.isCombinedCommonBilling === true;
    const combinedChildRecords = isCombinedCommonBilling
      ? Array.isArray(activeRecurringRecord.combinedChildRecords)
        ? activeRecurringRecord.combinedChildRecords
        : []
      : [];
    const normalizedSubscriptionType = normalizeValue(
      activeRecurringRecord.subscriptionType,
    ).toLowerCase();
    const isVariable = normalizedSubscriptionType === "variable";
    const monthlyAmount = normalizeValue(
      activeRecurringRecord?.monthlyBillingAmounts?.[recurringCycle.monthKey],
    );

    if (recurringDecision === "Hold") {
      const remarks = normalizeValue(recurringRemarks);

      if (!remarks) {
        alert("Remarks are required when recurring billing is placed on Hold.");
        return;
      }

      const savedAt = new Date().toISOString();
      const targetRecordIds = new Set(
        isCombinedCommonBilling
          ? combinedChildRecords.map((record) => record.recordId)
          : [activeRecurringRecord.recordId],
      );

      setStage3Records((currentRecords) =>
        currentRecords.map((record) =>
          targetRecordIds.has(record.recordId)
            ? {
                ...record,
                recurringBillingHolds: {
                  ...(record.recurringBillingHolds || {}),
                  [recurringCycle.monthKey]: {
                    status: "Hold",
                    remarks,
                    savedAt,
                  },
                },
              }
            : record,
        ),
      );

      alert(
        isCombinedCommonBilling
          ? "Combined complex recurring billing is on Hold for this billing cycle."
          : "Recurring billing is on Hold for this billing month.",
      );
      return;
    }

    if (isVariable && !monthlyAmount) {
      alert("Enter the current month billing amount before processing.");
      return;
    }

    const billingSubscriptionFee = isVariable
      ? monthlyAmount
      : normalizeValue(recurringCycle.billingSubscriptionFee);

    if (!billingSubscriptionFee) {
      alert("Recurring billing amount is required before processing.");
      return;
    }

    const confirmed = window.confirm(
      isCombinedCommonBilling
        ? `Process combined recurring billing for ${
            activeRecurringRecord.complexCode || "this complex"
          } for ${recurringCycle.billingMonth || "the current cycle"}?`
        : `Process recurring billing for ${
            recurringCycle.billingMonth || "the current month"
          }?`,
    );

    if (!confirmed) {
      return;
    }

    const submittedAt = new Date().toISOString();

    if (isCombinedCommonBilling) {
      if (combinedChildRecords.length === 0) {
        alert("Combined complex billing rows are not available.");
        return;
      }

      const groupKey =
        normalizeValue(activeRecurringRecord.combinedGroupKey) ||
        getCommonComplexBillingGroupKey(activeRecurringRecord);
      const commonSubscriptionFee = resolveCommonComplexSubscriptionFee(
        combinedChildRecords[0],
        stage3Records,
        commonComplexAllocations,
      );
      const childSnapshots = combinedChildRecords
        .map((childRecord) => {
          const childCycle = childRecord.recurringCycle || {};
          const childBillingAmount = normalizeValue(
            childCycle.billingSubscriptionFee,
          );

          if (!childBillingAmount) return null;

          const childRecordForSnapshot = {
            ...childRecord,
            billingPeriodFrom: childCycle.periodFrom,
            billingPeriodTo: childCycle.periodTo,
            subscriptionFee: childBillingAmount,
            billingVerificationStatus: "Submitted to Billing Team",
          };
          const childSnapshot = buildBillingRecordSnapshot(
            childRecordForSnapshot,
            stage3Records,
            commonComplexAllocations,
            submittedAt,
            "Submitted to Billing Team",
          );

          if (!childSnapshot) return null;

          childSnapshot.billingAmountBeforeGST = childBillingAmount;
          childSnapshot.billingSubscriptionFee = childBillingAmount;
          childSnapshot.actualSubscriptionFee = normalizeValue(
            childCycle.actualSubscriptionFee,
          );
          childSnapshot.billingMonth = normalizeValue(childCycle.billingMonth);
          childSnapshot.billingCycleType = "Recurring";
          childSnapshot.billingPeriodFrom = normalizeValue(childCycle.periodFrom);
          childSnapshot.billingPeriodTo = normalizeValue(childCycle.periodTo);
          childSnapshot.billingMode = "Combined Complex Billing - Site Breakup";
          childSnapshot.isCombinedChildMarker = true;
          childSnapshot.combinedGroupKey = groupKey;
          childSnapshot.billingRecordId = [
            "COMBINED-CHILD",
            groupKey,
            normalizeValue(childRecord.screenCode),
            normalizeValue(childCycle.periodFrom),
            normalizeValue(childCycle.periodTo),
          ].join("::");

          return childSnapshot;
        })
        .filter(Boolean);

      if (childSnapshots.length === 0) {
        alert("No billable site amount is available for the combined complex billing cycle.");
        return;
      }

      childSnapshots.forEach((snapshot) => {
        snapshot.currentCombinedInvoiceAmount = billingSubscriptionFee;
        snapshot.configuredCommonSubscriptionFee = normalizeValue(
          commonSubscriptionFee,
        );
        snapshot.combinedSiteCount = childSnapshots.length;
      });

      setBillingRecords((currentRecords) => {
        const nextRecords = [...currentRecords];

        childSnapshots.forEach((snapshot) => {
          const existingIndex = nextRecords.findIndex(
            (record) => record.billingRecordId === snapshot.billingRecordId,
          );

          if (existingIndex < 0) {
            nextRecords.push(snapshot);
          } else {
            nextRecords[existingIndex] = {
              ...nextRecords[existingIndex],
              ...snapshot,
            };
          }
        });

        return nextRecords;
      });

      const targetRecordIds = new Set(
        combinedChildRecords.map((record) => record.recordId),
      );

      setStage3Records((currentRecords) =>
        currentRecords.map((record) => {
          if (!targetRecordIds.has(record.recordId)) {
            return record;
          }

          const nextHolds = { ...(record.recurringBillingHolds || {}) };
          delete nextHolds[recurringCycle.monthKey];

          return {
            ...record,
            recurringBillingHolds: nextHolds,
            lastRecurringBillingProcessedAt: submittedAt,
            commonBillingMode: "Combined Complex Billing",
          };
        }),
      );

      setActiveRecordId("");
      setRecurringDecision("Process");
      setRecurringRemarks("");
      alert("Combined complex recurring billing processed and added to Billing Records.");
      return;
    }

    const recurringRecordForSnapshot = {
      ...activeRecurringRecord,
      billingPeriodFrom: recurringCycle.periodFrom,
      billingPeriodTo: recurringCycle.periodTo,
      subscriptionFee: billingSubscriptionFee,
      billingVerificationStatus: "Submitted to Billing Team",
    };

    const snapshot = buildBillingRecordSnapshot(
      recurringRecordForSnapshot,
      stage3Records,
      commonComplexAllocations,
      submittedAt,
      "Submitted to Billing Team",
    );

    if (!snapshot) {
      return;
    }

    snapshot.billingAmountBeforeGST = billingSubscriptionFee;
    snapshot.billingSubscriptionFee = billingSubscriptionFee;
    snapshot.actualSubscriptionFee = normalizeValue(
      recurringCycle.actualSubscriptionFee,
    );
    snapshot.billingMonth = normalizeValue(recurringCycle.billingMonth);
    snapshot.billingCycleType = "Recurring";
    snapshot.billingPeriodFrom = normalizeValue(recurringCycle.periodFrom);
    snapshot.billingPeriodTo = normalizeValue(recurringCycle.periodTo);
    snapshot.billingRecordId = getBillingRecordIdentity({
      screenCode: activeRecurringRecord.screenCode,
      billingPeriodFrom: recurringCycle.periodFrom,
      billingPeriodTo: recurringCycle.periodTo,
    });

    setBillingRecords((currentRecords) => {
      const existingIndex = currentRecords.findIndex(
        (record) => record.billingRecordId === snapshot.billingRecordId,
      );

      if (existingIndex < 0) {
        return [...currentRecords, snapshot];
      }

      const nextRecords = [...currentRecords];
      nextRecords[existingIndex] = {
        ...nextRecords[existingIndex],
        ...snapshot,
      };
      return nextRecords;
    });

    setStage3Records((currentRecords) =>
      currentRecords.map((record) => {
        if (record.recordId !== activeRecurringRecord.recordId) {
          return record;
        }

        const nextHolds = { ...(record.recurringBillingHolds || {}) };
        delete nextHolds[recurringCycle.monthKey];

        return {
          ...record,
          recurringBillingHolds: nextHolds,
          lastRecurringBillingProcessedAt: submittedAt,
        };
      }),
    );

    setActiveRecordId("");
    setRecurringDecision("Process");
    setRecurringRemarks("");
    alert("Recurring billing processed and added to Billing Records.");
  }

  function handleWorkspaceChange(nextWorkspace) {
    setActiveWorkspace(nextWorkspace);
  }

  function handleBillingSectionToggle(section) {
    setActiveBillingSection((currentSection) =>
      currentSection === section ? "" : section,
    );
  }


  useEffect(() => {
    let cancelled = false;

    async function restorePersistedExpenseSites() {
      try {
        const sites = await billingApiRequest("/sites");
        if (!cancelled) {
          setPersistedExpenseSites(Array.isArray(sites) ? sites : []);
        }
      } catch (error) {
        if (!cancelled) {
          console.error("Unable to restore persisted Other Expenses:", error);
        }
      }
    }

    restorePersistedExpenseSites();
    return () => {
      cancelled = true;
    };
  }, []);

  const otherExpenseRecords = useMemo(() => {
    const groupedExpenses = new Map();

    const recordsBySiteId = new Map();
    persistedExpenseSites.forEach((site) => {
      const stage1 = site?.stage1Data && typeof site.stage1Data === "object"
        ? site.stage1Data
        : {};
      const stage2 = site?.stage2Data && typeof site.stage2Data === "object"
        ? site.stage2Data
        : {};
      const stage3 = site?.stage3Data && typeof site.stage3Data === "object"
        ? site.stage3Data
        : {};
      const snapshot = { ...stage1, ...stage2, ...stage3 };
      const backendSiteId = normalizeValue(site?.id);
      const screenCode = normalizeValue(site?.siteId);

      if (!backendSiteId && !screenCode) return;

      recordsBySiteId.set(backendSiteId || screenCode, {
        ...snapshot,
        backendSiteId,
        billingCode: normalizeValue(site?.billingId),
        complexCode: normalizeValue(site?.complexId),
        screenCode,
        screenName: normalizeValue(site?.screenName),
        installationExpenses:
          stage3.installationExpenses ?? stage2.installationExpenses ?? [],
      });
    });

    // Live records take precedence for their own Site, but never replace a
    // sibling Site's persisted expense collection.
    [...stage2Records, ...stage3Records].forEach((record) => {
      const key = normalizeValue(record?.backendSiteId) || normalizeValue(record?.screenCode);
      if (key) recordsBySiteId.set(key, record);
    });

    Array.from(recordsBySiteId.values()).forEach((record) => {
      const billingCode = normalizeValue(record?.billingCode);
      const complexCode = normalizeValue(record?.complexCode);
      const screenCode = normalizeValue(record?.screenCode);
      const isComplex = Boolean(complexCode);

      const groupKey = isComplex
        ? `complex::${billingCode}::${complexCode}`
        : `standalone::${billingCode}::${screenCode || normalizeValue(record?.recordId)}`;

      if (!groupKey) {
        return;
      }

      const expenses = getOwnedInstallationExpenses(record);

      const applicable =
        normalizeValue(record?.otherInstallationExpensesApplicable) === "Yes" ||
        expenses.length > 0;

      if (!applicable || expenses.length === 0) {
        return;
      }

      const existing = groupedExpenses.get(groupKey) || {
        recordId: groupKey,
        billingCode,
        complexCode,
        screenCode: isComplex ? "" : screenCode,
        screenCodes: new Set(),
        displayName: isComplex
          ? normalizeValue(record?.complexName) ||
            normalizeValue(record?.billingName) ||
            complexCode
          : normalizeValue(record?.screenName),
        location: normalizeValue(record?.location),
        state: normalizeValue(record?.state),
        isComplex,
        expenseMap: new Map(),
      };

      if (screenCode) {
        existing.screenCodes.add(screenCode);
      }

      expenses.forEach((expense, index) => {
        const expenseKey =
          normalizeValue(expense?.expenseId) ||
          [
            normalizeValue(expense?.screenCode) || screenCode,
            normalizeValue(expense?.expenseType),
            normalizeValue(expense?.expenseDate),
            Number(expense?.amount || 0),
            normalizeValue(expense?.comments),
            index,
          ].join("::");

        if (!existing.expenseMap.has(expenseKey)) {
          existing.expenseMap.set(expenseKey, {
            ...expense,
            backendSiteId: normalizeValue(record?.backendSiteId),
            screenCode: normalizeValue(expense?.screenCode) || screenCode,
            screenName: normalizeValue(expense?.screenName) || normalizeValue(record?.screenName),
          });
        }
      });

      if (!existing.displayName) {
        existing.displayName = isComplex
          ? normalizeValue(record?.complexName) ||
            normalizeValue(record?.billingName) ||
            complexCode
          : normalizeValue(record?.screenName);
      }

      if (!existing.location) {
        existing.location = normalizeValue(record?.location);
      }

      if (!existing.state) {
        existing.state = normalizeValue(record?.state);
      }

      groupedExpenses.set(groupKey, existing);
    });

    return Array.from(groupedExpenses.values())
      .map((record) => {
        const expenses = Array.from(record.expenseMap.values());

        return {
          ...record,
          expenseMap: undefined,
          screenCodes: Array.from(record.screenCodes || []),
          expenses,
          total: expenses.reduce(
            (sum, expense) => sum + Number(expense?.amount || 0),
            0,
          ),
        };
      })
      .sort((left, right) =>
        (left.displayName || "").localeCompare(right.displayName || ""),
      );
  }, [persistedExpenseSites, stage2Records, stage3Records]);

  const filteredOtherExpenseRecords = useMemo(() => {
    const query = normalizeValue(otherExpensesSearchTerm).toLowerCase();
    const from = otherExpensesFromDate ? parseLocalDateValue(otherExpensesFromDate) : null;
    const to = otherExpensesToDate ? parseLocalDateValue(otherExpensesToDate) : null;

    return otherExpenseRecords
      .filter((record) => {
        if (query && ![
          record.billingCode,
          record.complexCode,
          record.screenCode,
          ...(record.screenCodes || []),
          record.displayName,
          record.location,
          record.state,
        ].join(" ").toLowerCase().includes(query)) return false;

        if (from || to) {
          const dates = (record.expenses || [])
            .map((expense) => parseLocalDateValue(normalizeValue(expense.expenseDate)))
            .filter(Boolean);
          if (!dates.length) return false;
          const matchesRange = dates.some((date) => (!from || date >= from) && (!to || date <= to));
          if (!matchesRange) return false;
        }
        return true;
      })
      .sort((left, right) => {
        const latest = (record) => Math.max(0, ...(record.expenses || []).map((expense) => parseLocalDateValue(normalizeValue(expense.expenseDate))?.getTime() || 0));
        const valueFor = (record) => {
          if (!otherExpensesSortKey) return latest(record);
          if (otherExpensesSortKey === "otherExpenses") return record.total || 0;
          return record?.[otherExpensesSortKey];
        };
        const diff = compareColumnValues(valueFor(left), valueFor(right));
        return otherExpensesSortOrder === "asc" ? diff : -diff;
      });
  }, [otherExpenseRecords, otherExpensesSearchTerm, otherExpensesFromDate, otherExpensesToDate, otherExpensesSortOrder, otherExpensesSortKey]);

  const otherExpensesPageCount = Math.max(1, Math.ceil(filteredOtherExpenseRecords.length / otherExpensesRowsPerPage));
  const pagedOtherExpenseRecords = useMemo(() => {
    const startIndex = (otherExpensesPage - 1) * otherExpensesRowsPerPage;
    return filteredOtherExpenseRecords.slice(startIndex, startIndex + otherExpensesRowsPerPage);
  }, [filteredOtherExpenseRecords, otherExpensesPage, otherExpensesRowsPerPage]);

  useEffect(() => {
    setOtherExpensesPage(1);
  }, [otherExpensesSearchTerm, otherExpensesFromDate, otherExpensesToDate, otherExpensesSortKey, otherExpensesSortOrder, otherExpensesRowsPerPage]);

  useEffect(() => {
    setOtherExpensesPage((currentPage) => Math.min(currentPage, otherExpensesPageCount));
  }, [otherExpensesPageCount]);

  const activeOtherExpenseRecord = useMemo(
    () =>
      otherExpenseRecords.find(
        (record) => record.recordId === activeOtherExpenseRecordId,
      ) || null,
    [otherExpenseRecords, activeOtherExpenseRecordId],
  );

  function handleOpenOtherExpenses(recordId) {
    setActiveOtherExpenseRecordId(recordId);
  }

  function handleOpenOtfTransaction(transactionId) {
    setActiveOtfTransactionId(transactionId);
    setActiveWorkspace(workspaceModes.OTF);
  }

  function updateActiveOtfTransaction(field, value) {
    if (!activeOtfTransaction) {
      return;
    }

    setOtfTransactionEdits((currentEdits) => ({
      ...currentEdits,
      [activeOtfTransaction.transactionId]: {
        ...(currentEdits[activeOtfTransaction.transactionId] || {}),
        [field]: value,
      },
    }));
  }

  function handleValidatedOtfDateChange(field) {
    return (event) => {
      updateActiveOtfTransaction(field, event.target.value);
    };
  }

  async function handleSaveOtfDetails() {
    if (!activeOtfTransaction) {
      return;
    }

    const invoiceNumber = normalizeValue(activeOtfTransaction.otfInvoiceNumber);
    const invoiceDate = normalizeValue(activeOtfTransaction.otfInvoiceDate);
    const receiptNumber = normalizeValue(activeOtfTransaction.otfRicbrNumber);
    const receiptDate = normalizeValue(
      activeOtfTransaction.otfRicbrCreatedDate,
    );
    const receivedAmount = normalizeValue(
      activeOtfTransaction.otfReceivedAmount,
    );

    if (!invoiceNumber) {
      alert("OTF Invoice Number is required.");
      return;
    }

    const invoiceConflict = findInvoiceConflict({
      candidateRecord: {
        ...activeOtfTransaction,
        invoiceTransactionGroup: activeOtfTransaction.transactionId,
      },
      candidateEntry: {
        invoiceNumber,
        periodFrom: activeOtfTransaction.otfInvoiceDate,
        periodTo: activeOtfTransaction.otfInvoiceDate,
      },
      billingRecords: [
        ...billingRecords,
        ...stage3Records
          .filter(
            (record) =>
              !(activeOtfTransaction.sourceRecordIds || []).includes(
                record.recordId,
              ),
          )
          .map((record) => ({
          ...record,
          billingPeriodFrom: record.otfInvoiceDate,
          billingPeriodTo: record.otfInvoiceDate,
          invoiceEntries: [
            {
              entryId: `otf-${record.recordId}`,
              invoiceNumber: record.otfInvoiceNumber,
              periodFrom: record.otfInvoiceDate,
              periodTo: record.otfInvoiceDate,
            },
          ],
          })),
      ],
    });

    if (invoiceConflict) {
      alert(invoiceConflict);
      return;
    }

    if (!invoiceDate) {
      alert("OTF Invoice Date is required.");
      return;
    }

    if (!isValidDateValue(invoiceDate)) {
      alert("Please enter a valid OTF Invoice Date.");
      return;
    }

    if (isFutureLocalDateValue(invoiceDate)) {
      alert("OTF Invoice Date must not be a future date.");
      return;
    }

    if (!receiptNumber) {
      alert("OTF Receipt / RICBR Number is required.");
      return;
    }

    if (!receiptDate) {
      alert("OTF Receipt / RICBR Date is required.");
      return;
    }

    if (!isValidDateValue(receiptDate)) {
      alert("Please enter a valid OTF Receipt / RICBR Date.");
      return;
    }

    if (isFutureLocalDateValue(receiptDate)) {
      alert("OTF Receipt / RICBR Date must not be a future date.");
      return;
    }

    const parsedReceivedAmount = parseAmountValue(receivedAmount);

    if (parsedReceivedAmount === null || parsedReceivedAmount <= 0) {
      alert("OTF Received Amount must be a valid amount greater than zero.");
      return;
    }

    const sourceRecordIds = new Set(
      (activeOtfTransaction.sourceRecordIds || [])
        .map((recordId) => normalizeValue(recordId))
        .filter(Boolean),
    );

    const includedScreenCodes = new Set(
      (activeOtfTransaction.includedSiteIds || [])
        .map((screenCode) => normalizeValue(screenCode))
        .filter(Boolean),
    );

    const linkedRecords = combinedCommercialRecords.filter((record) =>
      sourceRecordIds.has(normalizeValue(record.recordId)) ||
      includedScreenCodes.has(normalizeValue(record.screenCode)),
    );
    const linkedSites = Array.from(
      new Map(
        linkedRecords
          .filter((record) => normalizeValue(record.backendSiteId))
          .map((record) => [normalizeValue(record.backendSiteId), record]),
      ).values(),
    );

    if (linkedSites.length === 0) {
      alert("Unable to persist OTF details because no canonical Site records are linked.");
      return;
    }

    const updatedAt = new Date().toISOString();
    const persistedOtfFields = {
      otfApplicable: "Yes",
      otfType: normalizeValue(activeOtfTransaction.otfType),
      otfAmount: normalizeValue(activeOtfTransaction.otfAmount),
      otfInvoiceNumber: invoiceNumber,
      otfInvoiceDate: invoiceDate,
      otfRicbrNumber: receiptNumber,
      otfRicbrCreatedDate: receiptDate,
      otfReceiptNumber: receiptNumber,
      otfReceiptDate: receiptDate,
      otfReceivedAmount: formatAmountValue(parsedReceivedAmount),
      otfPaymentReceivedAmount: formatAmountValue(parsedReceivedAmount),
      otfPaymentStatus: "Paid",
      otfPaymentReceivedDate: receiptDate,
      otfPaymentReference: receiptNumber,
      otfIncludedScreenCodes: Array.from(
        new Set([
          ...includedScreenCodes,
          ...linkedRecords.map((record) => normalizeValue(record.screenCode)),
        ].filter(Boolean)),
      ),
      otfTransactionGroupKey: activeOtfTransaction.transactionId,
      otfDetailsUpdatedAt: updatedAt,
    };

    try {
      await Promise.all(
        linkedSites.map(async (record) => {
          const persistedSite = await billingApiRequest(
            `/sites/${normalizeValue(record.backendSiteId)}`,
          );
          const existingStage3Data =
            persistedSite?.stage3Data &&
            typeof persistedSite.stage3Data === "object" &&
            !Array.isArray(persistedSite.stage3Data)
              ? persistedSite.stage3Data
              : {};

          await billingApiRequest(
            `/sites/${normalizeValue(record.backendSiteId)}`,
            {
              method: "PATCH",
              body: JSON.stringify({
                stage3Data: {
                  ...existingStage3Data,
                  ...persistedOtfFields,
                },
              }),
            },
          );
        }),
      );
    } catch (error) {
      alert(
        error instanceof Error
          ? `Unable to save OTF details. ${error.message}`
          : "Unable to save OTF details.",
      );
      return;
    }

    const applyPersistedOtfFields = (currentRecords) =>
      currentRecords.map((record) => {
        const sameSource =
          sourceRecordIds.size > 0 &&
          sourceRecordIds.has(normalizeValue(record.recordId));
        const sameScreen =
          includedScreenCodes.size > 0 &&
          includedScreenCodes.has(normalizeValue(record.screenCode));

        if (!sameSource && !sameScreen) {
          return record;
        }

        return { ...record, ...persistedOtfFields };
      });
    setStage2Records(applyPersistedOtfFields);
    setStage3Records(applyPersistedOtfFields);

    setOtfTransactionEdits((currentEdits) => {
      const nextEdits = { ...currentEdits };
      delete nextEdits[activeOtfTransaction.transactionId];
      return nextEdits;
    });

    setActiveOtfTransactionId("");
    alert("OTF details saved and linked to Incentive eligibility.");
  }

  function handleCombinedPaymentDownload() {
    const exportRows = buildCombinedPaymentDownloadRows(
      otfTransactions,
      stage3Records,
    );

    if (exportRows.length === 0) {
      return;
    }

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Payment Download");
    XLSX.writeFile(
      workbook,
      `stage3-payment-download-${new Date().toISOString().slice(0, 10)}.xlsx`,
    );
  }

  function handleDownloadFirstTimeBillingExcel() {
    if (firstTimeBillingDownloadRows.length === 0) {
      alert("No completed First Time Billing records are available for download.");
      return;
    }

    const exportRows = buildFirstTimeBillingTeamDownloadRows(
      firstTimeBillingDownloadRows,
      stage3Records,
      commonComplexAllocations,
    );

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "3A First Time Billing");
    XLSX.writeFile(
      workbook,
      `3A_First_Time_Billing_${getDownloadDateTimeStamp()}.xlsx`,
    );
  }

  function handleDownloadRecurringBillingExcel() {
    if (recurringBillingDownloadRows.length === 0) {
      alert("No 3B Recurring Billing records are available for download.");
      return;
    }

    const blockedRows = recurringBillingDownloadRows.filter(
      (record) => record.recurringCycle?.priceChangeBlocked === true,
    );
    let eligibleRows = recurringBillingDownloadRows;

    if (blockedRows.length > 0) {
      const blockedLabels = Array.from(
        new Set(
          blockedRows.map(
            (record) =>
              record.recurringCycle?.priceChangeBlockLabel ||
              record.screenName ||
              record.screenCode ||
              record.complexCode,
          ),
        ),
      ).filter(Boolean);

      const proceed = window.confirm(
        `${blockedLabels.join(", ")} ${
          blockedLabels.length === 1 ? "is" : "are"
        } currently under Price Change & Reallocation. Recurring Billing cannot be processed using the previous commercial value.\n\nTo proceed with the remaining eligible sites/complexes while excluding the blocked item(s), select Yes/OK. Select No/Cancel to go back and validate Price Change & Reallocation.`,
      );

      if (!proceed) return;

      eligibleRows = recurringBillingDownloadRows.filter(
        (record) => record.recurringCycle?.priceChangeBlocked !== true,
      );
    }

    if (eligibleRows.length === 0) {
      alert("No eligible Recurring Billing records remain after excluding Price Change & Reallocation items.");
      return;
    }

    const exportRows = eligibleRows.map((record) =>
      buildRecurringBillingTeamExportRow(record),
    );

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "3B Recurring Billing");
    XLSX.writeFile(
      workbook,
      `3B_Recurring_Billing_${getDownloadDateTimeStamp()}.xlsx`,
    );
  }

  function toggleVerificationSelection(recordId, checked) {
    const record = selectableVerificationRows.find(
      (row) => row.recordId === recordId,
    );

    if (!record) {
      return;
    }

    setSelectedVerificationIds((currentIds) =>
      checked
        ? Array.from(new Set([...currentIds, recordId]))
        : currentIds.filter((id) => id !== recordId),
    );
  }

  function handleToggleAllVerificationSelections(checked) {
    const visibleRecordIds = pagedSelectableVerificationRows.map(
      (record) => record.recordId,
    );

    setSelectedVerificationIds((currentIds) =>
      checked
        ? Array.from(new Set([...currentIds, ...visibleRecordIds]))
        : currentIds.filter((id) => !visibleRecordIds.includes(id)),
    );
  }

  async function handleVerifyAndSendToBillingTeam() {
    const selectedRecords = selectableVerificationRows.filter((record) =>
      selectedVerificationIds.includes(record.recordId),
    );

    if (selectedRecords.length === 0) {
      alert(
        "Open Billing Information & Invoice and complete the required site billing details before Verify & Submit.",
      );
      return;
    }

    if (
      selectedRecords.some(
        (record) => getPersistedFirstTimeBillingStatus(record) === "Completed",
      )
    ) {
      alert("Completed First Time Billing records are already submitted.");
      return;
    }

    const activeRecordIsSelected = selectedRecords.some(
      (record) => record.recordId === activeRecordId,
    );
    const waiverValidationError = activeRecordIsSelected
      ? getFirstTimeWaiverValidationError(
          firstTimeWaiverDraft,
          firstTimeMonthlyBillingRows,
          completedActivePeriodKeys,
          activeRecord?.billingStartDate,
        )
      : "";

    if (waiverValidationError) {
      setFirstTimeValidationMessage(waiverValidationError);
      return;
    }

    const firstValidationError = selectedRecords
      .flatMap((record) => [
        ...getBillingValidationErrors(record),
        getCommonComplexBillingValidationError(
          record,
          stage3Records,
          commonComplexAllocations,
        ),
      ])
      .find((message) => Boolean(message));

    if (firstValidationError) {
      alert(firstValidationError);
      return;
    }

    const confirmed = window.confirm(
      "Please confirm the Subscription Fee, billing period, payment details, and other billing information have been verified.\n\nOnce submitted, this billing record will be marked Completed and locked for further editing.",
    );

    if (!confirmed) {
      return;
    }

    const selectedRecordIds = new Set(
      selectedRecords.map((record) => record.recordId),
    );
    const submittedAt = new Date().toISOString();
    const billingSnapshots = selectedRecords
      .map((record) => {
        const recordForSnapshot =
          record.recordId === activeRecordId
            ? {
                ...record,
                firstTimeBillingValidationStatus: "Completed",
                firstTimeBillingCompletedAt: submittedAt,
                firstTimeBillingCompletedBy: "Operations",
                firstTimeWaiver: {
                  ...firstTimeWaiverDraft,
                  reason: normalizeValue(firstTimeWaiverDraft.reason),
                },
                firstTimePeriodBreakdown,
                firstTimeBillingSummary,
              }
            : record;
        const snapshot = buildBillingRecordSnapshot(
          recordForSnapshot,
          stage3Records,
          commonComplexAllocations,
          submittedAt,
          "Submitted to Billing Team",
        );

        return snapshot
          ? {
              ...snapshot,
              billingCycleType: "First Time",
              billingLifecycleStatus:
                normalizeValue(snapshot.billingLifecycleStatus) || "Active",
            }
          : null;
      })
      .filter(Boolean);

    const approvedWaiver = activeRecordIsSelected
      ? {
          ...firstTimeWaiverDraft,
          reason: normalizeValue(firstTimeWaiverDraft.reason),
          savedAt: submittedAt,
        savedBy: "Operations",
      }
      : null;
    const submittedFirstTimeBillingRecord =
      billingSnapshots.find(
        (snapshot) =>
          getFirstTimeBillingSiteKey(snapshot) ===
          getFirstTimeBillingSiteKey(activeRecord),
      ) || null;

    const completedStage3Record = activeRecordIsSelected
      ? {
          ...activeRecord,
          firstTimeBillingValidationStatus: "Completed",
          firstTimeBillingCompletedAt: submittedAt,
          firstTimeBillingCompletedBy: "Operations",
          billingVerificationStatus: "Submitted to Billing Team",
          billingVerifiedAt: submittedAt,
          firstTimeWaiver: approvedWaiver,
          firstTimePeriodBreakdown,
          firstTimeBillingSummary,
          firstTimeBillingRecord: submittedFirstTimeBillingRecord,
        }
      : null;
    const persistedCompletedRecords = new Map();

    try {
      for (const record of selectedRecords) {
        const backendSiteId = normalizeValue(record?.backendSiteId);
        if (!backendSiteId) {
          throw new Error(
            `Unable to persist Completed status for ${record?.screenCode || "the selected site"}: backend Site ID is missing.`,
          );
        }

        const snapshot = billingSnapshots.find(
          (candidate) =>
            getFirstTimeBillingSiteKey(candidate) ===
            getFirstTimeBillingSiteKey(record),
        );
        const completedRecord =
          record.recordId === activeRecordId
            ? completedStage3Record
            : {
                ...record,
                firstTimeBillingValidationStatus: "Completed",
                firstTimeBillingCompletedAt: submittedAt,
                firstTimeBillingCompletedBy: "Operations",
                billingVerificationStatus: "Submitted to Billing Team",
                billingVerifiedAt: submittedAt,
                firstTimeBillingRecord: snapshot || null,
              };
        const savedSite = await billingApiRequest(`/sites/${backendSiteId}`, {
          method: "PATCH",
          body: JSON.stringify({
            stage3Data: {
              ...completedRecord,
              ...(record.recordId === activeRecordId
                ? {
                    firstTimeWaiver: approvedWaiver,
                    firstTimePeriodBreakdown,
                    firstTimeBillingSummary,
                  }
                : {}),
            },
          }),
        });
        const persistedStage3Data =
          savedSite?.stage3Data &&
          typeof savedSite.stage3Data === "object" &&
          !Array.isArray(savedSite.stage3Data)
            ? savedSite.stage3Data
            : {};
        persistedCompletedRecords.set(record.recordId, {
          ...completedRecord,
          ...persistedStage3Data,
          backendSiteId: savedSite?.id || backendSiteId,
        });
      }
    } catch (error) {
      setFirstTimeValidationMessage(
        error instanceof Error ? error.message : "Unable to persist Completed status.",
      );
      return;
    }

    setStage3Records((currentRecords) =>
      currentRecords.map((record) =>
        selectedRecordIds.has(record.recordId)
          ? {
              ...record,
              ...(persistedCompletedRecords.get(record.recordId) || {}),
              billingVerificationStatus: "Submitted to Billing Team",
              billingVerifiedAt: submittedAt,
              billingCommercialLocked: true,
              billingCommercialLockedAt: submittedAt,
              billingCommercialLockReason: "Data shared with Billing Team",
              billingCommercialLockSnapshot: {
                installationDate: normalizeValue(record.installationDate),
                liveDate: normalizeValue(record.liveDate),
                trialPeriod: normalizeValue(record.trialPeriod),
                trialPeriodExtension: normalizeValue(
                  record.trialPeriodExtension ||
                    record.totalTrialPeriodExtension ||
                    record.trialExtension,
                ),
                billingStartDate: normalizeValue(record.billingStartDate),
                mouStartDate: normalizeValue(record.mouStartDate),
                mouEndDate: normalizeValue(record.mouEndDate),
                subscriptionType: normalizeValue(record.subscriptionType),
                subscriptionMode: normalizeValue(record.subscriptionMode),
                subscriptionFee: normalizeValue(record.subscriptionFee),
                otfApplicable: normalizeValue(record.otfApplicable),
                otfType: normalizeValue(record.otfType),
                otfAmount: normalizeValue(record.otfAmount),
                pricingMethod: normalizeValue(record.pricingMethod),
                pricingGroups: Array.isArray(record.pricingGroups)
                  ? record.pricingGroups.map((group) => ({ ...group }))
                  : [],
              },
              billingLifecycleStatus:
                normalizeValue(record.billingLifecycleStatus) || "Active",
              ...(record.recordId === activeRecordId && approvedWaiver
                ? {
                    firstTimeWaiver: approvedWaiver,
                    firstTimePeriodBreakdown,
                    firstTimeBillingSummary,
                  }
                : {}),
              ...(normalizeValue(record.correctionType) === "Human Error" &&
              normalizeValue(record.correctionRequestedAt)
                ? {
                    correctionReverifiedAt: submittedAt,
                    correctionType: "",
                    correctionRoute: "",
                  }
                : {}),
            }
          : record,
      ),
    );

    if (selectedRecordIds.has(activeRecordId)) {
      setActiveRecordId("");
    }

    setBillingRecords((currentRecords) => {
      const nextRecords = [...currentRecords];

      billingSnapshots.forEach((snapshot) => {
        const firstTimeSiteKey = getFirstTimeBillingSiteKey(snapshot);

        const existingFirstTimeIndex = nextRecords.findIndex(
          (record) =>
            !isRecurringBillingRecord(record) &&
            getFirstTimeBillingSiteKey(record) === firstTimeSiteKey,
        );

        // A completed first-time Billing Record is permanent history.
        // Returning the site through Stage 2/Stage 1 must not create or overwrite it.
        if (existingFirstTimeIndex >= 0) {
          const selectedSourceRecord = selectedRecords.find(
            (record) =>
              getFirstTimeBillingSiteKey(record) === firstTimeSiteKey,
          );

          if (
            selectedSourceRecord &&
            normalizeValue(selectedSourceRecord.correctionType) === "Human Error" &&
            normalizeValue(selectedSourceRecord.correctionRequestedAt)
          ) {
            const existingRecord = nextRecords[existingFirstTimeIndex];
            const reverifyAuditEntry = createBillingAuditEntry({
              event: "Record Correction Re-verified",
              field: "billingVerificationStatus",
              previousValue: "Correction Pending Re-verification",
              newValue: "Submitted to Billing Team",
              remarks:
                normalizeValue(selectedSourceRecord.correctionReason) ||
                "Corrected site re-verified in Stage 3.",
              timestamp: submittedAt,
            });

            nextRecords[existingFirstTimeIndex] = {
              ...existingRecord,
              lastUpdatedAt: submittedAt,
              auditTrail: [
                ...(Array.isArray(existingRecord.auditTrail)
                  ? existingRecord.auditTrail
                  : []),
                reverifyAuditEntry,
              ],
            };
          }

          return;
        }

        nextRecords.push(snapshot);
      });

      return nextRecords;
    });


    setSelectedVerificationIds((currentIds) =>
      currentIds.filter((recordId) => !selectedRecordIds.has(recordId)),
    );

    lockCommonComplexAllocations(selectedRecords);

    alert(
      "Billing record submitted to Billing Team.",
    );
  }

  function handleOpenHumanErrorCorrection(billingRecordId) {
    const selectedBillingRecord = billingRecords.find(
      (record) => record.billingRecordId === billingRecordId,
    );

    if (!selectedBillingRecord) {
      return;
    }

    setHumanErrorCorrectionRecordId(selectedBillingRecord.billingRecordId);
    setHumanErrorCorrectionFieldsSelected([]);
    setHumanErrorCorrectionRemarks("");
  }

  function handleCloseHumanErrorCorrection() {
    setHumanErrorCorrectionRecordId("");
    setHumanErrorCorrectionFieldsSelected([]);
    setHumanErrorCorrectionRemarks("");
  }

  function handleToggleHumanErrorCorrectionField(fieldKey, checked) {
    setHumanErrorCorrectionFieldsSelected((currentFields) =>
      checked
        ? Array.from(new Set([...currentFields, fieldKey]))
        : currentFields.filter((key) => key !== fieldKey),
    );
  }

  async function handleSubmitHumanErrorCorrection() {
    if (!humanErrorCorrectionRecord) {
      return;
    }

    if (humanErrorCorrectionFieldsSelected.length === 0) {
      alert("Select at least one field to correct.");
      return;
    }

    const remarks = normalizeValue(humanErrorCorrectionRemarks);

    if (!remarks) {
      alert("Correction Reason / Remarks is required.");
      return;
    }

    const requestedAt = new Date().toISOString();
    const selectedFields = [...humanErrorCorrectionFieldsSelected];
    const correctionEventId = crypto.randomUUID();
    const correctionRecord = {
      ...humanErrorCorrectionRecord,
      humanErrorCorrection: {
        type: "Human Error",
        requestedAt,
        remarks,
        fields: selectedFields,
        requiresBillingDateReview: selectedFields.includes("subscriptionMode"),
        sourceBillingRecordId:
          humanErrorCorrectionBillingRecord?.billingRecordId || "",
      },
      correctionRequestedAt: requestedAt,
      correctionReason: remarks,
      correctionType: "Human Error",
      correctionFields: selectedFields,
      requiresBillingDateReview: selectedFields.includes("subscriptionMode"),
      workflowType: "RECORD_CORRECTION",
      correctionEventId,
    };

    let correctionAuditId = "";
    if (!humanErrorCorrectionRecord?.backendSiteId) {
      alert("Unable to persist the Record Correction request because the backend site reference is missing.");
      return;
    }

    if (humanErrorCorrectionRecord.backendSiteId) {
      try {
        const audit = await billingApiRequest(
          `/sites/${humanErrorCorrectionRecord.backendSiteId}/corrections`,
          {
            method: "POST",
            body: JSON.stringify({
              workflowType: "RECORD_CORRECTION",
              eventId: correctionEventId,
              correctionFields: selectedFields,
              reason: remarks,
              previousValues: selectedFields.reduce((values, fieldKey) => {
                values[fieldKey] = humanErrorCorrectionRecord[fieldKey] ?? null;
                return values;
              }, {}),
              newValues: {},
              requestedAt,
            }),
          },
        );
        correctionAuditId = audit?.id || "";
      } catch (error) {
        alert(`Unable to persist the Record Correction request. ${error.message}`);
        return;
      }
    }

    correctionRecord.correctionAuditId = correctionAuditId;

    if (humanErrorCorrectionBillingRecord?.billingRecordId) {
      const correctionAuditEntry = createBillingAuditEntry({
        event: "Record Correction Requested",
        field: "correctionFields",
        previousValue: "-",
        newValue: selectedFields.map(
          (fieldKey) =>
            humanErrorCorrectionFields.find((field) => field.key === fieldKey)
              ?.label || fieldKey,
        ),
        remarks,
        timestamp: requestedAt,
      });

      setBillingRecords((currentRecords) =>
        currentRecords.map((record) =>
          record.billingRecordId === humanErrorCorrectionBillingRecord.billingRecordId
            ? {
                ...record,
                auditTrail: [
                  ...(Array.isArray(record.auditTrail) ? record.auditTrail : []),
                  correctionAuditEntry,
                ],
                lastUpdatedAt: requestedAt,
              }
            : record,
        ),
      );
    }

    const movedSuccessfully =
      await onReturnCommercialCorrectionToStage1(correctionRecord);

    if (!movedSuccessfully) {
      alert("Unable to send the selected record to Stage 1 for correction.");
      return;
    }

    handleCloseHumanErrorCorrection();

    alert(
      selectedFields.includes("subscriptionMode")
        ? "Record Correction requested. The selected site has moved to Stage 1 for MoU & Subscription correction. Subscription Mode change will require Billing Date review in Stage 2 before returning to Stage 3."
        : "Record Correction requested. The selected site has moved to Stage 1 for MoU & Subscription correction.",
    );
  }

  function appendStageMovementAudit(record, destinationStage, requestedAt) {
    if (!record || !destinationStage) {
      return;
    }

    const sourceRecordId = normalizeValue(record.recordId);
    const screenCode = normalizeValue(record.screenCode);

    const matchingBillingRecord = billingRecords
      .filter((billingRecord) => {
        const sameSource =
          sourceRecordId &&
          normalizeValue(billingRecord.sourceRecordId) === sourceRecordId;
        const sameScreen =
          screenCode &&
          normalizeValue(billingRecord.screenCode || billingRecord.siteScope) ===
            screenCode;

        return sameSource || sameScreen;
      })
      .sort((left, right) => {
        const leftTime = new Date(
          left.submittedAt || left.createdAt || left.lastUpdatedAt || 0,
        ).getTime();
        const rightTime = new Date(
          right.submittedAt || right.createdAt || right.lastUpdatedAt || 0,
        ).getTime();

        return rightTime - leftTime;
      })[0];

    if (!matchingBillingRecord) {
      return;
    }

    const stageAuditEntry = createBillingAuditEntry({
      event: "Human Error Stage Movement",
      field: "stage",
      previousValue: "Stage 3",
      newValue: destinationStage,
      remarks:
        normalizeValue(record.correctionReason) ||
        `Moved to ${destinationStage} for correction.`,
      timestamp: requestedAt,
    });

    setBillingRecords((currentRecords) =>
      currentRecords.map((billingRecord) =>
        billingRecord.billingRecordId === matchingBillingRecord.billingRecordId
          ? {
              ...billingRecord,
              auditTrail: [
                ...(Array.isArray(billingRecord.auditTrail)
                  ? billingRecord.auditTrail
                  : []),
                stageAuditEntry,
              ],
              lastUpdatedAt: requestedAt,
            }
          : billingRecord,
      ),
    );
  }

  async function handleFirstTimeBillingRowAction(record, action) {
    if (!record || !action) {
      return;
    }

    if (action === "Open") {
      handleOpenVerificationRecord(record.recordId);
      return;
    }

    if (!isPendingFirstTimeBillingRecord(record)) {
      alert(
        "First Time Billing is already processed. Later commercial changes must be handled through Billing Records → Action → Price Change & Reallocation.",
      );
      return;
    }

    const requestedAt = new Date().toISOString();
    const correctionRecord = {
      ...record,
      firstTimeBillingValidatedAt: "",
      correctionType: "Human Error",
      correctionRequestedAt: requestedAt,
      correctionRoute: action,
      humanErrorCorrection: {
        ...(record.humanErrorCorrection || {}),
        type: "Human Error",
        requestedAt,
        route: action,
      },
    };

    if (action === "Stage 1") {
      appendStageMovementAudit(record, "Stage 1", requestedAt);

      const movedSuccessfully =
        await onReturnCommercialCorrectionToStage1(correctionRecord);

      if (!movedSuccessfully) {
        alert("Unable to move the selected site to Stage 1.");
      }

      return;
    }

    if (action === "Stage 2") {
      appendStageMovementAudit(record, "Stage 2", requestedAt);

      const movedSuccessfully = await onReturnToStage2([correctionRecord]);

      if (!movedSuccessfully) {
        alert("Unable to move the selected site to Stage 2.");
      }
    }
  }

  async function handleReturnToStage2() {
    if (!activeRecord) {
      return;
    }

    const correctionRecord = {
      ...activeRecord,
      firstTimeBillingValidatedAt: "",
    };

    const movedSuccessfully = await onReturnToStage2([correctionRecord]);
    if (!movedSuccessfully) {
      alert("Unable to return the selected record to Stage 2.");
    }
  }

  function handleOpenClosureRecord(recordId) {
    const record = closureRows.find(
      (row) =>
        row.recordId === recordId ||
        row.billingRecordId === recordId ||
        row.lifecycleSourceRecordId === recordId,
    );

    if (!record) {
      return;
    }

    setClosureEditorMode("reason");
    setActiveClosureRecordId(recordId);
    setClosureDraft({
      closureType:
        record.currentBillingStatus === "Inactive"
          ? ""
          : record.currentBillingStatus,
      effectiveDate: normalizeValue(record.closureEffectiveDate),
      remarks: normalizeValue(record.closureRemarks),
    });
    setRestoreActiveDraft({
      activeFrom: "",
      remarks: "",
    });
    setClosureMessage("");

    window.requestAnimationFrame(() => {
      const element = document.getElementById("stage3-billing-closure-editor");
      element?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }

  function handleOpenReactivate(recordId) {
    const record = closureRows.find(
      (row) =>
        row.recordId === recordId ||
        row.billingRecordId === recordId ||
        row.lifecycleSourceRecordId === recordId,
    );

    if (!record) {
      return;
    }

    if (
      !["Billing Paused", "Site Inactive"].includes(
        normalizeValue(record.currentBillingStatus),
      )
    ) {
      return;
    }

    setClosureEditorMode("reactivate");
    setActiveClosureRecordId(recordId);
    setRestoreActiveDraft({
      activeFrom: "",
      remarks: "",
    });
    setClosureMessage("");

    window.requestAnimationFrame(() => {
      const element = document.getElementById("stage3-billing-closure-editor");
      element?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }

  function updateClosureDraft(field, value) {
    setClosureDraft((currentDraft) => ({
      ...currentDraft,
      [field]: value,
    }));
    setClosureMessage("");
  }

  function handleSaveBillingClosure() {
    if (!activeClosureRecord) {
      return;
    }

    const closureType = normalizeValue(closureDraft.closureType);
    const effectiveDate = normalizeValue(closureDraft.effectiveDate);
    const remarks = normalizeValue(closureDraft.remarks);

    if (
      ![
        "Billing Paused",
        "Site Inactive",
        "Site Removed",
        "Contract Closed",
      ].includes(closureType)
    ) {
      setClosureMessage("Please select a valid Closure Type.");
      return;
    }

    if (!effectiveDate) {
      setClosureMessage("Effective Date is required.");
      return;
    }

    if (!isValidDateValue(effectiveDate)) {
      setClosureMessage("Please enter a valid Effective Date.");
      return;
    }

    if (!remarks) {
      setClosureMessage("Remarks / Reason is required.");
      return;
    }

    const now = new Date().toISOString();

    const activeClosureSourceRecordId = normalizeValue(
      activeClosureRecord.lifecycleSourceRecordId ||
        activeClosureRecord.sourceRecordId ||
        activeClosureRecord.recordId,
    );
    const activeClosureScreenCode = normalizeValue(
      activeClosureRecord.screenCode || activeClosureRecord.siteScope,
    );

    setStage3Records((currentRecords) =>
      currentRecords.map((record) => {
        const sameSource =
          activeClosureSourceRecordId &&
          normalizeValue(record.recordId) === activeClosureSourceRecordId;
        const sameScreen =
          activeClosureScreenCode &&
          normalizeValue(record.screenCode) === activeClosureScreenCode;

        return sameSource || sameScreen
          ? {
              ...record,
              billingLifecycleStatus: "Inactive",
              billingLifecycleReason: closureType,
              closureEffectiveDate: effectiveDate,
              pauseFromDate:
                closureType === "Billing Paused"
                  ? effectiveDate
                  : record.pauseFromDate || "",
              closureRemarks: remarks,
              closureUpdatedAt: now,
              closureHistory: [
                ...(Array.isArray(record.closureHistory)
                  ? record.closureHistory
                  : []),
                {
                  eventId: `closure-${Date.now()}-${Math.random()
                    .toString(36)
                    .slice(2, 7)}`,
                  previousStatus:
                    normalizeValue(record.billingLifecycleReason) || "Active",
                  newStatus: closureType,
                  effectiveDate,
                  remarks,
                  updatedAt: now,
                },
              ],
            }
          : record;
      }),
    );

    const sourceRecordId = normalizeValue(activeClosureRecord.recordId);
    const screenCode = normalizeValue(activeClosureRecord.screenCode);

    setBillingRecords((currentRecords) =>
      currentRecords.map((record) => {
        const sameBillingRecord =
          normalizeValue(activeClosureRecord.billingRecordId) &&
          normalizeValue(record.billingRecordId) ===
            normalizeValue(activeClosureRecord.billingRecordId);
        const sameSource =
          sourceRecordId &&
          normalizeValue(record.sourceRecordId) === sourceRecordId;
        const sameScreen =
          screenCode &&
          normalizeValue(record.screenCode || record.siteScope) === screenCode;

        if (!sameBillingRecord && !sameSource && !sameScreen) {
          return record;
        }

        return {
          ...record,
          billingLifecycleStatus: "Inactive",
          billingStatusReason: closureType,
          closureEffectiveDate: effectiveDate,
          pauseFromDate:
            closureType === "Billing Paused"
              ? effectiveDate
              : record.pauseFromDate || "",
          lastUpdatedAt: now,
          auditTrail: [
            ...(Array.isArray(record.auditTrail) ? record.auditTrail : []),
            createBillingAuditEntry({
              event: "Paused / Inactive Reason Updated",
              field: "billingStatusReason",
              previousValue:
                normalizeValue(record.billingStatusReason) || "Inactive",
              newValue: closureType,
              remarks,
              timestamp: now,
            }),
          ],
        };
      }),
    );

    setClosureMessage(`${closureType} saved successfully.`);
  }

  function handleRestoreToActive(recordId) {
    const record = closureRows.find(
      (row) =>
        row.recordId === recordId ||
        row.billingRecordId === recordId ||
        row.lifecycleSourceRecordId === recordId,
    );

    if (!record) {
      return;
    }

    const currentReason = normalizeValue(record.currentBillingStatus);

    if (
      ["Site Removed", "Contract Closed"].includes(currentReason)
    ) {
      setClosureMessage(
        `${currentReason === "Site Removed" ? "Site Removed" : "Contract Quit / Closed"} cannot be reactivated from Stage 3. If the site rejoins, the workflow must start from the beginning.`,
      );
      return;
    }

    if (!["Billing Paused", "Site Inactive"].includes(currentReason)) {
      setClosureMessage(
        "Only Billing Paused or Site Inactive records can be reactivated from this screen.",
      );
      return;
    }

    const activeFrom = normalizeValue(restoreActiveDraft.activeFrom);
    const restoreRemarks = normalizeValue(restoreActiveDraft.remarks);

    if (!activeFrom) {
      setClosureMessage("Active From date is required.");
      return;
    }

    if (!isValidDateValue(activeFrom)) {
      setClosureMessage("Please enter a valid Active From date.");
      return;
    }

    const pausedFromDate = getBillingPauseDate(record);
    const minimumActiveFrom = pausedFromDate
      ? formatLocalDateValue(pausedFromDate)
      : "";

    if (minimumActiveFrom && activeFrom < minimumActiveFrom) {
      setClosureMessage("Active From cannot be earlier than Paused From.");
      return;
    }

    const confirmed = window.confirm(
      `Restore ${record.screenName || record.screenCode} to Active billing from ${formatBillingDateDisplay(activeFrom)}?`,
    );

    if (!confirmed) {
      return;
    }

    const now = new Date().toISOString();
    const targetScreenCode = normalizeValue(
      record.screenCode || record.siteScope,
    );
    const targetSourceRecordId = normalizeValue(
      record.lifecycleSourceRecordId ||
        record.sourceRecordId ||
        record.recordId,
    );

    setStage3Records((currentRecords) =>
      currentRecords.map((item) => {
        const sameSource =
          targetSourceRecordId &&
          normalizeValue(item.recordId) === targetSourceRecordId;
        const sameScreen =
          targetScreenCode &&
          normalizeValue(item.screenCode) === targetScreenCode;

        if (!sameSource && !sameScreen) {
          return item;
        }

        return {
          ...item,
          billingLifecycleStatus: "Active",
          billingLifecycleReason: "",
          billingActiveFromDate: activeFrom,
          closureEffectiveDate: "",
          pauseFromDate: "",
          closureRemarks: "",
          inactiveSince: "",
          inactiveEffectiveDate: "",
          closureUpdatedAt: now,
          closureHistory: [
            ...(Array.isArray(item.closureHistory)
              ? item.closureHistory
              : []),
            {
              eventId: `restore-active-${Date.now()}-${Math.random()
                .toString(36)
                .slice(2, 7)}`,
              previousStatus: currentReason,
              newStatus: "Active",
              effectiveDate: activeFrom,
              remarks:
                restoreRemarks ||
                `Restored to Active from ${currentReason}.`,
              updatedAt: now,
            },
          ],
        };
      }),
    );

    setBillingRecords((currentRecords) =>
      currentRecords.map((billingRecord) => {
        const sameSource =
          targetSourceRecordId &&
          normalizeValue(billingRecord.sourceRecordId) ===
            targetSourceRecordId;
        const sameScreen =
          targetScreenCode &&
          normalizeValue(
            billingRecord.screenCode || billingRecord.siteScope,
          ) === targetScreenCode;

        if (!sameSource && !sameScreen) {
          return billingRecord;
        }

        return {
          ...billingRecord,
          billingLifecycleStatus: "Active",
          billingStatusReason: "",
          billingActiveFromDate: activeFrom,
          inactiveSince: "",
          inactiveEffectiveDate: "",
          closureEffectiveDate: "",
          lastUpdatedAt: now,
          auditTrail: [
            ...(Array.isArray(billingRecord.auditTrail)
              ? billingRecord.auditTrail
              : []),
            createBillingAuditEntry({
              event: "Billing Restored to Active",
              field: "billingStatus",
              previousValue: currentReason,
              newValue: "Active",
              remarks:
                restoreRemarks ||
                `Active From ${formatBillingDateDisplay(activeFrom)}.`,
              timestamp: now,
            }),
          ],
        };
      }),
    );

    setRestoreActiveDraft({
      activeFrom: "",
      remarks: "",
    });
    setClosureEditorMode("reason");
    setActiveClosureRecordId("");
    setClosureDraft({
      closureType: "",
      effectiveDate: "",
      remarks: "",
    });
    setClosureMessage("Site restored to Active successfully.");
  }

  async function handleValidateFirstTimeBillingSite() {
    if (!activeRecord || firstTimeValidationInFlightRef.current) {
      return;
    }

    if (isFirstTimeBillingSiteValidatedForCurrentCycle(activeRecord)) {
      setFirstTimeValidationMessage(
        "This site is already Validated and Ready for Verify & Submit.",
      );
      return;
    }

    const validationErrors = getBillingValidationErrors(activeRecord);
    const commonValidationError = getCommonComplexBillingValidationError(
      activeRecord,
      stage3Records,
      commonComplexAllocations,
    );

    if (validationErrors.length > 0) {
      setFirstTimeValidationMessage(validationErrors[0]);
      return;
    }

    if (commonValidationError) {
      setFirstTimeValidationMessage(commonValidationError);
      return;
    }

    const normalizedSubscriptionType = normalizeValue(
      activeRecord.subscriptionType,
    ).toLowerCase();
    const normalizedSubscriptionMode = normalizeValue(
      activeRecord.subscriptionMode,
    )
      .toLowerCase()
      .replace(/[-_]/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    const isVariableMonthly =
      normalizedSubscriptionType === "variable" &&
      ["", "monthly", "month"].includes(normalizedSubscriptionMode);

    if (isVariableMonthly) {
      const eligibleRows = firstTimeMonthlyBillingRows.filter(
        (row) => row.invoiceEligible === true,
      );
      const missingRow = eligibleRows.find(
        (row) => !normalizeValue(row.amount),
      );

      if (eligibleRows.length === 0) {
        setFirstTimeValidationMessage(
          "No eligible completed billing month is available for validation.",
        );
        return;
      }

      if (missingRow) {
        setFirstTimeValidationMessage(
          `${missingRow.billingMonth}: Billing amount is required before validation.`,
        );
        return;
      }
    } else if (
      !isFirstTimeBillingTeamDownloadEligible(
        activeRecord,
        stage3Records,
        commonComplexAllocations,
      )
    ) {
      setFirstTimeValidationMessage(
        "Complete the required first-time billing information before validation.",
      );
      return;
    }

    const backendSiteId = normalizeValue(activeRecord.backendSiteId);

    if (!backendSiteId) {
      setFirstTimeValidationMessage(
        "Unable to persist validation because the backend Site ID is missing.",
      );
      return;
    }

    const validatedAt = new Date().toISOString();
    const validatedBy = "Operations";
    const validatedRecord = {
      ...activeRecord,
      firstTimeBillingValidationStatus: "Validated",
      firstTimeBillingValidatedAt: validatedAt,
      firstTimeBillingValidatedBy: validatedBy,
      firstTimeWaiver: {
        ...firstTimeWaiverDraft,
        reason: normalizeValue(firstTimeWaiverDraft.reason),
      },
      firstTimePeriodBreakdown,
      firstTimeBillingSummary,
    };
    let persistedRecord = validatedRecord;

    firstTimeValidationInFlightRef.current = true;

    try {
      const savedSite = await billingApiRequest(`/sites/${backendSiteId}`, {
        method: "PATCH",
        body: JSON.stringify({ stage3Data: validatedRecord }),
      });
      const persistedStage3Data =
        savedSite?.stage3Data &&
        typeof savedSite.stage3Data === "object" &&
        !Array.isArray(savedSite.stage3Data)
          ? savedSite.stage3Data
          : validatedRecord;

      persistedRecord = {
        ...validatedRecord,
        ...persistedStage3Data,
        backendSiteId: savedSite?.id || backendSiteId,
      };
    } catch (error) {
      setFirstTimeValidationMessage(
        error instanceof Error
          ? error.message
          : "Unable to persist First Time Billing validation.",
      );
      return;
    } finally {
      firstTimeValidationInFlightRef.current = false;
    }

    setStage3Records((currentRecords) =>
      currentRecords.map((record) =>
        record.recordId === activeRecord.recordId
          ? {
              ...record,
              ...persistedRecord,
            }
          : record,
      ),
    );

    setFirstTimeValidationMessage(
      "Validated successfully. This site is Ready for Verify & Submit.",
    );

    window.requestAnimationFrame(() => {
      guidedScrollToElement(firstTimeBillingRef.current);
    });
  }

  // Shared first-time billing workspace.
  // Keep this rendered only under 3A; 3B recurring billing has its own workflow.
  function renderBillingInformationAndInvoice() {
    const showsBillingAllocation = Boolean(
      commonComplexBillingSummary &&
        !isVariableSubscriptionRecord(activeRecord),
    );
    const calculationSectionNumber = showsBillingAllocation ? 2 : 1;
    const waiverSectionNumber = calculationSectionNumber + 1;
    const summarySectionNumber = calculationSectionNumber + 2;
    const verificationSectionNumber = calculationSectionNumber + 3;
    const billablePeriodCount = firstTimePeriodBreakdown.filter(
      (row) => row.treatment === "Billable",
    ).length;
    const waivedPeriodCount = firstTimePeriodBreakdown.filter(
      (row) => row.treatment === "Waived",
    ).length;

    return (
      <article
        className="billings-page__card"
        ref={billingInformationRef}
      >
          <div className="billings-page__card-header">
            <div>
              <p className="billings-page__card-title">Billing Information &amp; Invoice</p>
              <span className="billings-page__card-badge">Linked</span>
            </div>
          </div>

          <p className="billings-page__helper">
            Review the selected Stage 3 billing record. Complete the required
            billing details in the sections below, then validate the site before
            using Verify &amp; Submit in the First Time Billing list.
          </p>

          <div className="billings-page__field-grid">
            <div className="billings-page__field">
              <label>Billing Code / Customer Code</label>
              <input type="text" value={activeRecord?.billingCode || ""} readOnly />
            </div>

            <div className="billings-page__field">
              <label>Complex Code</label>
              <input
                type="text"
                value={activeRecord?.complexCode || "Standalone"}
                readOnly
              />
            </div>

            <div className="billings-page__field">
              <label>Screen Code</label>
              <input type="text" value={activeRecord?.screenCode || ""} readOnly />
            </div>

            <div className="billings-page__field">
              <label>Screen Name</label>
              <input type="text" value={activeRecord?.screenName || ""} readOnly />
            </div>

            <div className="billings-page__field">
              <label>Location</label>
              <input type="text" value={activeRecord?.location || ""} readOnly />
            </div>

            <div className="billings-page__field">
              <label>Current Stage Status</label>
              <input
                type="text"
                value={activeRecord?.currentStageStatus || ""}
                readOnly
              />
            </div>

            {getCorrectionDisplayLabel(activeRecord) ? (
              <>
                <div className="billings-page__field">
                  <label>Workflow Type</label>
                  <input
                    type="text"
                    value={getCorrectionDisplayLabel(activeRecord)}
                    readOnly
                  />
                </div>

                <div className="billings-page__field">
                  <label>Correction Fields</label>
                  <input
                    type="text"
                    value={
                      Array.isArray(activeRecord?.correctionFields)
                        ? activeRecord.correctionFields.join(", ")
                        : ""
                    }
                    readOnly
                  />
                </div>
              </>
            ) : null}
          </div>

          <div className="billings-page__billing-section-heading">
            <h2>Billing &amp; Invoice Details</h2>
            <span className="billings-page__billing-section-text">
              Selected site billing information is shown below in the full-width workspace.
            </span>
          </div>

          <div
            className={`billings-page__billing-form${
              showsBillingAllocation
                ? " billings-page__billing-form--with-allocation"
                : ""
            }`}
          >
            <section className="billings-page__billing-section billings-page__first-time-section--calculation">
              <div className="billings-page__billing-section-heading">
                <h2>{calculationSectionNumber}. Billing Calculation</h2>
                <span className="billings-page__billing-section-text">
                  Auto-calculated based on Billing Start Date, subscription terms
                  and applicable billing rules.
                </span>
              </div>

              <div className="billings-page__billing-section-grid">
                <div className="billings-page__field">
                  <label>Billing Start Date</label>
                  <input
                    type="date"
                    value={activeRecord?.billingStartDate || ""}
                    readOnly
                    disabled={!activeRecord}
                  />
                </div>

                <div className="billings-page__field">
                  <label>Subscription Type</label>
                  <input
                    type="text"
                    value={activeRecord?.subscriptionType || ""}
                    readOnly
                  />
                </div>

                <div className="billings-page__field">
                  <label>Subscription Mode</label>
                  <input
                    type="text"
                    value={activeRecord?.subscriptionMode || ""}
                    readOnly
                  />
                </div>
              </div>

              {activeRecord && isMonthlyBilling ? (
                <>
                  <div className="billings-page__verification-table-wrapper" style={{ maxHeight: 360, overflowY: "auto" }}>
                  <table className="billings-page__verification-table">
                    <thead>
                      <tr>
                        <th>Billing Month</th>
                        <th>Period From</th>
                        <th>Period To</th>
                        <th>Subscription Type</th>
                        <th>Subscription Fee</th>
                        <th>Treatment</th>
                        <th>Billable Amount</th>
                        <th>Invoice Eligibility</th>
                      </tr>
                    </thead>

                    <tbody>
                      {firstTimePeriodBreakdown.map((row) => {
                        const isVariable =
                          normalizeValue(row.subscriptionType).toLowerCase() ===
                          "variable";

                        return (
                          <tr key={row.monthKey}>
                            <td>{row.billingMonth}</td>
                            <td>{row.periodFrom}</td>
                            <td>{row.periodTo}</td>
                            <td>{row.subscriptionType || "-"}</td>
                            <td>
                              {isVariable ? (
                                <input
                                  type="text"
                                  inputMode="decimal"
                                  value={row.amount}
                                  onChange={(event) =>
                                    handleVariableMonthlyAmountChange(
                                      row.monthKey,
                                      event.target.value,
                                    )
                                  }
                                  placeholder="Enter monthly amount"
                                  style={{
                                    width: "100%",
                                    minWidth: 0,
                                    height: 36,
                                    boxSizing: "border-box",
                                  }}
                                />
                              ) : (
                                row.amount || "-"
                              )}
                            </td>
                            <td>{row.treatment}</td>
                            <td>{row.billableAmount || "0"}</td>
                            <td>{row.invoiceEligibilityLabel}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
                </>
              ) : null}
            </section>

            {activeRecord && isMonthlyBilling ? (
              <section className="billings-page__billing-section billings-page__first-time-section--waiver">
                <div className="billings-page__billing-section-heading">
                  <h2>{waiverSectionNumber}. Waiver</h2>
                  <span className="billings-page__billing-section-text">
                    Apply a waiver only when an approved billing-period waiver is required.
                  </span>
                </div>

                <div className="billings-page__billing-section-grid">
                  <div className="billings-page__field">
                    <label htmlFor="first-time-waiver-applicable">Waiver Applicable *</label>
                    <select
                      id="first-time-waiver-applicable"
                      value={firstTimeWaiverDraft.applicable}
                      onChange={(event) =>
                        setFirstTimeWaiverDraft((current) => ({
                          ...current,
                          applicable: event.target.value,
                          ...(event.target.value === "No"
                            ? { from: "", to: "", reason: "" }
                            : {}),
                        }))
                      }
                    >
                      <option value="No">No</option>
                      <option value="Yes">Yes</option>
                    </select>
                  </div>

                  {firstTimeWaiverDraft.applicable === "Yes" ? (
                    <>
                      <div className="billings-page__field">
                        <label htmlFor="first-time-waiver-from">Waiver From *</label>
                        <input
                          id="first-time-waiver-from"
                          type="date"
                          value={firstTimeWaiverDraft.from}
                          onChange={(event) =>
                            setFirstTimeWaiverDraft((current) => ({
                              ...current,
                              from: event.target.value,
                            }))
                          }
                        />
                      </div>
                      <div className="billings-page__field">
                        <label htmlFor="first-time-waiver-to">Waiver To *</label>
                        <input
                          id="first-time-waiver-to"
                          type="date"
                          value={firstTimeWaiverDraft.to}
                          onChange={(event) =>
                            setFirstTimeWaiverDraft((current) => ({
                              ...current,
                              to: event.target.value,
                            }))
                          }
                        />
                      </div>
                      <div className="billings-page__field">
                        <label htmlFor="first-time-waiver-reason">Waiver Reason *</label>
                        <input
                          id="first-time-waiver-reason"
                          type="text"
                          value={firstTimeWaiverDraft.reason}
                          onChange={(event) =>
                            setFirstTimeWaiverDraft((current) => ({
                              ...current,
                              reason: event.target.value,
                            }))
                          }
                        />
                      </div>
                    </>
                  ) : null}
                </div>
              </section>
            ) : null}

            {activeRecord && isMonthlyBilling ? (
              <section className="billings-page__billing-section billings-page__first-time-section--summary">
                <div className="billings-page__billing-section-heading">
                  <h2>{summarySectionNumber}. Billing Summary</h2>
                </div>

                <div className="billings-page__first-time-summary">
                  <div className="billings-page__first-time-summary-item">
                    <span>Calculated Billing Amount</span>
                    <strong>{formatInrAmount(firstTimeBillingSummary.totalBeforeWaiver)}</strong>
                  </div>
                  <div className="billings-page__first-time-summary-item">
                    <span>Waiver Amount</span>
                    <strong>{formatInrAmount(firstTimeBillingSummary.waiverAmount)}</strong>
                  </div>
                  <div className="billings-page__first-time-summary-item billings-page__first-time-summary-item--actual">
                    <span>Actual Billing Amount</span>
                    <strong>{formatInrAmount(firstTimeBillingSummary.actualBillingAmount)}</strong>
                  </div>
                </div>
              </section>
            ) : null}

            {commonComplexBillingSummary &&
              !isVariableSubscriptionRecord(activeRecord) && (
              <section
                className="billings-page__billing-section billings-page__billing-section--common-allocation billings-page__first-time-section--allocation"
                ref={commonComplexAllocationRef}
              >
                <div className="billings-page__billing-section-heading">
                  <h2>1. Billing Allocation</h2>
                  <span className="billings-page__billing-section-text">
                    Review and save the screen-wise allocation for this Common Pricing
                    complex before validating billing.
                  </span>
                </div>

                <div className="billings-page__billing-allocation-summary">
                  <div className="billings-page__field">
                    <label>Billing Code / Customer Code</label>
                    <input type="text" value={commonComplexBillingSummary.billingCode} readOnly />
                  </div>

                  <div className="billings-page__field">
                    <label>Complex Code</label>
                    <input
                      type="text"
                      value={getDisplayComplexCode(commonComplexBillingSummary.complexCode)}
                      readOnly
                    />
                  </div>

                  <div className="billings-page__field">
                    <label>Billing Mode</label>
                    <input
                      type="text"
                      value={commonComplexBillingSummary.billingMode || "-"}
                      readOnly
                    />
                  </div>

                  <div className="billings-page__field">
                    <label>Current Subscription Fee</label>
                    <input
                      type="text"
                      value={commonComplexBillingSummary.commonSubscriptionFee}
                      readOnly
                    />
                  </div>

                  <div className="billings-page__field">
                    <label>Total Complex Sites</label>
                    <input
                      type="text"
                      value={commonComplexBillingSummary.totalComplexSites}
                      readOnly
                    />
                  </div>

                  <div className="billings-page__field">
                    <label>Currently Billable Sites</label>
                    <input
                      type="text"
                      value={commonComplexBillingSummary.currentlyBillableSites}
                      readOnly
                    />
                  </div>

                  <div className="billings-page__field">
                    <label>Pending Sites</label>
                    <input
                      type="text"
                      value={commonComplexBillingSummary.pendingSites}
                      readOnly
                    />
                  </div>

                  <div className="billings-page__field">
                    <label>Allocation Total</label>
                    <input
                      type="text"
                      value={commonComplexBillingSummary.allocationTotal}
                      readOnly
                    />
                  </div>

                  <div className="billings-page__field">
                    <label>Difference</label>
                    <input
                      type="text"
                      value={commonComplexBillingSummary.difference}
                      readOnly
                    />
                  </div>
                </div>

                {commonComplexBillingValidationError && (
                  <div className="billings-page__validation">
                    <p>{commonComplexBillingValidationError}</p>
                  </div>
                )}

                <div className="billings-page__billing-allocation-table-wrapper">
                  <table className="billings-page__billing-allocation-table">
                    <colgroup>
                      <col className="billings-page__billing-allocation-col-site" />
                      <col className="billings-page__billing-allocation-col-screen" />
                      <col className="billings-page__billing-allocation-col-start" />
                      <col className="billings-page__billing-allocation-col-status" />
                      <col className="billings-page__billing-allocation-col-allocation" />
                      <col className="billings-page__billing-allocation-col-charge" />
                      <col className="billings-page__billing-allocation-col-action" />
                    </colgroup>
                    <thead>
                      <tr>
                        <th>Screen Code</th>
                        <th>Screen / Complex Name</th>
                        <th>Billing Start Date</th>
                        <th>Billing Status</th>
                        <th>Allocated Fee</th>
                        <th>Current Billing Charge</th>
                        <th className="billings-page__foc-action-header">Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {commonComplexBillingRows.map((row) => {
                        const isSelectedComplexSite =
                          activeRecord &&
                          row.siteKey ===
                            getCommonComplexBillingSiteKey(activeRecord.screenCode) &&
                          !isFirstTimeBillingSiteValidatedForCurrentCycle(
                            activeRecord,
                          );

                        return (
                        <tr
                          key={row.siteKey}
                          style={
                            isSelectedComplexSite
                              ? {
                                  backgroundColor: "#17324a",
                                  boxShadow: "inset 4px 0 0 #f0b429",
                                }
                              : undefined
                          }
                        >
                          <td
                            title={row.screenCode || "-"}
                            style={
                              isSelectedComplexSite
                                ? { backgroundColor: "#17324a" }
                                : undefined
                            }
                          >
                            {row.screenCode || "-"}
                          </td>
                          <td
                            title={row.screenName || "-"}
                            style={
                              isSelectedComplexSite
                                ? { backgroundColor: "#17324a" }
                                : undefined
                            }
                          >
                            {row.screenName || "-"}
                          </td>
                          <td
                            title={row.billingStartDate || "-"}
                            style={
                              isSelectedComplexSite
                                ? { backgroundColor: "#17324a" }
                                : undefined
                            }
                          >
                            {row.billingStartDate || "-"}
                          </td>
                          <td
                            style={
                              isSelectedComplexSite
                                ? { backgroundColor: "#17324a" }
                                : undefined
                            }
                          >
                            {row.billingStatus}
                          </td>
                          <td
                            style={
                              isSelectedComplexSite
                                ? { backgroundColor: "#17324a" }
                                : undefined
                            }
                          >
                            <input
                              type="text"
                              inputMode="decimal"
                              value={row.allocatedFee || ""}
                              onChange={(event) =>
                                handleCommonComplexAllocationInlineChange(
                                  row.siteKey,
                                  event.target.value,
                                )
                              }
                              readOnly={row.locked}
                              placeholder="Enter allocation"
                              aria-label={`Allocated fee for ${row.screenCode}`}
                              style={{ width: "100%", minWidth: 0, boxSizing: "border-box" }}
                            />
                          </td>
                          <td
                            title={row.currentBillingCharge || "-"}
                            style={
                              isSelectedComplexSite
                                ? { backgroundColor: "#17324a" }
                                : undefined
                            }
                          >
                            {row.currentBillingCharge || "-"}
                          </td>
                          <td
                            className="billings-page__billing-allocation-table-action-cell"
                            style={
                              isSelectedComplexSite
                                ? { backgroundColor: "#17324a" }
                                : undefined
                            }
                          >
                            <span>{row.locked ? "Locked" : "Editable"}</span>
                          </td>
                        </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

                <div
                  className="billings-page__detail-actions"
                  style={{ justifyContent: "flex-end", marginTop: "14px" }}
                >
                  <button
                    type="button"
                    className="billings-page__primary-button"
                    onClick={handleSaveAllCommonComplexAllocations}
                  >
                    Save Complex Allocation
                  </button>
                </div>
              </section>
            )}

            <section className="billings-page__billing-section billings-page__first-time-section--supporting">
              <div className="billings-page__billing-section-heading">
                <h2>OTF Details</h2>
                <span className="billings-page__billing-section-text">
                  OTF commercial values remain read-only from Stage 1B.
                </span>
              </div>

              <div className="billings-page__billing-section-grid">
                <div className="billings-page__field">
                  <label>OTF Applicable</label>
                  <input
                    type="text"
                    value={
                      isOtfApplicable ? "Yes" : "No / Not Applicable"
                    }
                    readOnly
                  />
                </div>

                <div className="billings-page__field">
                  <label>OTF Type</label>
                  <input type="text" value={activeRecord?.otfType || ""} readOnly />
                </div>

                <div className="billings-page__field">
                  <label>OTF Amount Before GST</label>
                  <input type="text" value={activeRecord?.otfAmount || ""} readOnly />
                </div>

                <div className="billings-page__field">
                  <label>OTF Invoice Number</label>
                  <input
                    type="text"
                    value={isOtfApplicable ? activeRecord?.otfInvoiceNumber || "" : ""}
                    onChange={(event) =>
                      updateActiveRecord("otfInvoiceNumber", event.target.value)
                    }
                    disabled={!activeRecord || !isOtfApplicable}
                    placeholder={
                      isOtfApplicable ? "OTF invoice number" : "Not applicable"
                    }
                  />
                </div>
              </div>
            </section>

            <section className="billings-page__billing-section billings-page__first-time-section--verification">
              <div className="billings-page__billing-section-heading">
                <h2>{verificationSectionNumber}. Verification &amp; Submit</h2>
                <span className="billings-page__billing-section-text">
                  Validate this site, then select its Ready row in First Time Billing
                  and use the existing Verify &amp; Submit action.
                </span>
              </div>

              {activeRecord && isMonthlyBilling ? (
                <div className="billings-page__first-time-confirmation">
                  <span>Billing Periods: <strong>{firstTimePeriodBreakdown.length}</strong></span>
                  <span>Billable Periods: <strong>{billablePeriodCount}</strong></span>
                  <span>Waived Periods: <strong>{waivedPeriodCount}</strong></span>
                  <span>
                    Actual Billing Amount:{" "}
                    <strong>{formatInrAmount(firstTimeBillingSummary.actualBillingAmount)}</strong>
                  </span>
                </div>
              ) : null}

              <div className="billings-page__field">
                <label>Billing Remarks</label>
                <textarea
                  rows={3}
                  value={activeRecord?.billingRemarks || ""}
                  onChange={(event) =>
                    updateActiveRecord("billingRemarks", event.target.value)
                  }
                  disabled={!activeRecord}
                  placeholder="Enter billing remarks"
                />
              </div>

              {firstTimeValidationMessage ? (
                <div className="billings-page__validation">
                  <p>{firstTimeValidationMessage}</p>
                </div>
              ) : null}

              {activeValidationErrors.length > 0 && (
                <div className="billings-page__validation">
                  {activeValidationErrors.map((message) => (
                    <p key={message}>{message}</p>
                  ))}
                </div>
              )}

              <div className="billings-page__detail-actions billings-page__first-time-validation-actions">
                <button
                  type="button"
                  className="billings-page__primary-button"
                  onClick={handleValidateFirstTimeBillingSite}
                  disabled={!activeRecord}
                >
                  {activeRecord &&
                  isFirstTimeBillingSiteValidatedForCurrentCycle(activeRecord)
                    ? "Validated ✓"
                    : "Validate"}
                </button>
              </div>
            </section>
          </div>
      </article>
    );
  }

  return (
    <section className="billings-page">
      {billingSummaryError && (
        <div className="billings-page__modal-notice" role="alert">
          {billingSummaryError} Existing local data is being preserved.
        </div>
      )}
      <div
        className="billings-page__header-controls"
        aria-label="Stage 3 controls"
        style={{
          display: "flex",
          justifyContent: "flex-end",
          alignItems: "center",
          gap: 8,
          flexWrap: "wrap",
          width: "100%",
          marginBottom: 10,
        }}
      >
        <div
          className="billings-page__workspace-toggle"
          role="tablist"
          aria-label="Stage 3 workspace"
        >
          <button
            type="button"
            role="tab"
            className={
              activeWorkspace === workspaceModes.EXPENSES
                ? "billings-page__workspace-toggle-button billings-page__workspace-toggle-button--active"
                : "billings-page__workspace-toggle-button"
            }
            aria-selected={activeWorkspace === workspaceModes.EXPENSES}
            onClick={() => handleWorkspaceChange(workspaceModes.EXPENSES)}
          >
            Other Expenses
          </button>

          <button
            type="button"
            role="tab"
            className={
              activeWorkspace === workspaceModes.OTF
                ? "billings-page__workspace-toggle-button billings-page__workspace-toggle-button--active"
                : "billings-page__workspace-toggle-button"
            }
            aria-selected={activeWorkspace === workspaceModes.OTF}
            onClick={() => handleWorkspaceChange(workspaceModes.OTF)}
          >
            OTF
          </button>

          <button
            type="button"
            role="tab"
            className={
              activeWorkspace === workspaceModes.BILLING
                ? "billings-page__workspace-toggle-button billings-page__workspace-toggle-button--active"
                : "billings-page__workspace-toggle-button"
            }
            aria-selected={activeWorkspace === workspaceModes.BILLING}
            onClick={() => handleWorkspaceChange(workspaceModes.BILLING)}
          >
            Billing
          </button>

          <button
            type="button"
            role="tab"
            className={
              activeWorkspace === workspaceModes.FOC
                ? "billings-page__workspace-toggle-button billings-page__workspace-toggle-button--active"
                : "billings-page__workspace-toggle-button"
            }
            aria-selected={activeWorkspace === workspaceModes.FOC}
            onClick={() => handleWorkspaceChange(workspaceModes.FOC)}
          >
            FoC
          </button>
        </div>

        <button
          type="button"
          className="billings-page__primary-button billings-page__primary-button--compact"
          onClick={handleCombinedPaymentDownload}
          disabled={!hasCombinedPaymentRows}
        >
          Payment Download
        </button>
      </div>

      {activeWorkspace === workspaceModes.BILLING ? (
        <>
      <article
        className="billings-page__card billings-page__verification-card"
        ref={firstTimeBillingRef}
      >
        <div
          style={{
            width: "100%",
            marginBottom: 10,
          }}
        >
          <div
            className="billings-page__card-header"
            style={{
              display: "flex",
              alignItems: "center",
              gap: 12,
              width: "100%",
              padding: "6px 0 8px",
              boxSizing: "border-box",
            }}
          >
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <p
                className="billings-page__card-title"
                style={{ margin: 0, whiteSpace: "nowrap" }}
              >
                First Time Billing
              </p>
              <span className="billings-page__card-badge billings-page__card-badge--editable">
                {billingSummaryLoading
                  ? "Loading..."
                  : `${formatSummaryCount(canonicalBillingSummary?.firstTimeBilling)} Records`}
              </span>
            </div>

            <div
              style={{
                marginLeft: "auto",
                padding: 3,
                border: "1px solid rgba(96, 165, 250, 0.28)",
                borderRadius: 12,
                flexShrink: 0,
              }}
            >
              <button
                type="button"
                className="billings-page__secondary-button"
                style={{ minWidth: 105 }}
                onClick={() =>
                  handleBillingSectionToggle(billingSections.FIRST_TIME)
                }
              >
                {activeBillingSection === billingSections.FIRST_TIME
                  ? "Minimize"
                  : "Expand"}
              </button>
            </div>
          </div>

        </div>

        {activeBillingSection === billingSections.FIRST_TIME && (
          <>
            <div
              className="billings-page__verification-actions"
              style={{
                display: "grid",
                gridTemplateColumns: "minmax(0, 1fr) 150px 150px 150px 170px 180px",
                columnGap: 12,
                rowGap: 10,
                width: "100%",
                boxSizing: "border-box",
                alignItems: "end",
              }}
            >

              <div
                className="billings-page__field"
                style={{
                  marginBottom: 0,
                  minWidth: 0,
                  gridColumn: "1",
                }}
              >
                <label htmlFor="stage3-first-time-billing-search">Search</label>
                <input
                  id="stage3-first-time-billing-search"
                  type="search"
                  value={searchTerm}
                  onChange={(event) => setSearchTerm(event.target.value)}
                  placeholder="Search Billing Code / Customer Code, Complex Code, Screen Code, Screen Name, Billing Name, or Location"
                  style={{
                    width: "100%",
                    boxSizing: "border-box",
                    minWidth: 0,
                  }}
                />
              </div>

              <div
                className="billings-page__field"
                style={{
                  marginBottom: 0,
                  minWidth: 0,
                  gridColumn: "2",
                }}
              >
                <label htmlFor="stage3-first-time-status-filter">Status</label>
                <select
                  id="stage3-first-time-status-filter"
                  value={firstTimeStatusFilter}
                  onChange={(event) =>
                    setFirstTimeStatusFilter(event.target.value)
                  }
                  style={{
                    width: "100%",
                    boxSizing: "border-box",
                    minWidth: 0,
                  }}
                >
                  <option value="All">All</option>
                  <option value="Pending">Pending</option>
                  <option value="Validated">Validated</option>
                  <option value="Completed">Completed</option>
                </select>
              </div>

              <div className="billings-page__field" style={{ marginBottom: 0, minWidth: 0, gridColumn: "3" }}>
                <label htmlFor="stage3-first-time-from-date">From Date</label>
                <input id="stage3-first-time-from-date" type="date" value={firstTimeFromDate} onChange={(event) => setFirstTimeFromDate(event.target.value)} />
              </div>

              <div className="billings-page__field" style={{ marginBottom: 0, minWidth: 0, gridColumn: "4" }}>
                <label htmlFor="stage3-first-time-to-date">To Date</label>
                <input id="stage3-first-time-to-date" type="date" value={firstTimeToDate} onChange={(event) => setFirstTimeToDate(event.target.value)} />
              </div>

              <button
                type="button"
                className="billings-page__secondary-button"
                onClick={handleDownloadFirstTimeBillingExcel}
                disabled={firstTimeBillingDownloadRows.length === 0}
                style={{
                  gridColumn: "5",
                  width: "100%",
                  whiteSpace: "nowrap",
                }}
              >
                Download Excel
              </button>

              <button
                type="button"
                className="billings-page__primary-button"
                onClick={handleVerifyAndSendToBillingTeam}
                disabled={!hasSelectedVerificationRows}
                style={{
                  gridColumn: "6",
                  width: "100%",
                  whiteSpace: "nowrap",
                }}
              >
                Verify &amp; Submit
              </button>
            </div>

            {firstTimeVisibleRows.length === 0 ? (
              <div className="billings-page__verification-empty">
                No Stage 3 billing verification records are available yet.
              </div>
            ) : (
              <>
                <div className="billings-page__verification-table-wrapper">
                  <table
                    className="billings-page__verification-table"
                    style={{ width: "100%", tableLayout: "fixed" }}
                  >
                    <colgroup>
                      <col className="billings-page__verification-col-select" style={{ width: "52px" }} />
                      <col className="billings-page__verification-col-site" style={{ width: "112px" }} />
                      <col className="billings-page__verification-col-screen" />
                      <col className="billings-page__verification-col-billing-date" style={{ width: "112px" }} />
                      <col className="billings-page__verification-col-action" style={{ width: "130px" }} />
                      <col className="billings-page__verification-col-action" style={{ width: "120px" }} />
                      <col className="billings-page__verification-col-action" style={{ width: "90px" }} />
                    </colgroup>
                    <thead>
                      <tr>
                        <th className="billings-page__verification-table-select-header">
                          <input
                            ref={verificationSelectAllRef}
                            type="checkbox"
                            checked={areAllVisibleVerificationRowsSelected}
                            onChange={(event) =>
                              handleToggleAllVerificationSelections(event.target.checked)
                            }
                            aria-label="Select all billing verification records on this page"
                            title="Select all on this page"
                            disabled={!pagedSelectableVerificationRows.length}
                          />
                        </th>
                        <th style={{ paddingRight: 8 }}>Screen Code</th>
                        <th style={{ paddingRight: 8 }}>Screen Name</th>
                        <th style={{ paddingRight: 8 }}>Billing Date</th>
                        <th style={{ paddingRight: 8 }}>Subscription Fee</th>
                        <th style={{ paddingRight: 8 }}>Status</th>
                        <th
                          className="billings-page__verification-table-action-header"
                          style={{ paddingLeft: 6 }}
                        >
                          Action
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {pagedVerificationRows.map((record) => {
                        const isSelected = selectedVerificationIds.includes(record.recordId);
                        const isPending = isPendingFirstTimeBillingRecord(record);
                        const isValidated =
                          getPersistedFirstTimeBillingStatus(record) === "Validated" &&
                          isPending &&
                          isFirstTimeBillingSiteValidatedForCurrentCycle(record);
                        const isSelectable =
                          isFirstTimeBillingSiteReadyForSubmission(
                            record,
                            stage3Records,
                            commonComplexAllocations,
                          );
                        const isFirstBillingCompletedThisMonth =
                          getPersistedFirstTimeBillingStatus(record) === "Completed" ||
                          (!isSelectable &&
                            isFirstBillingSubmittedInCurrentMonth(
                              record,
                              billingRecords,
                            ));
                        const isBillingCompleted =
                          isFirstBillingCompletedThisMonth ||
                          isFirstTimeBillingTeamDownloadEligible(
                            record,
                            stage3Records,
                            commonComplexAllocations,
                          );

                        const isCurrentOpenSite =
                          record.recordId === activeRecordId &&
                          !isFirstTimeBillingSiteValidatedForCurrentCycle(record);

                        return (
                          <tr
                            key={record.recordId}
                            title={
                              isFirstBillingCompletedThisMonth
                                ? "First Billing Completed / Moves to 3B Next Month"
                                : isBillingCompleted
                                  ? "Completed"
                                  : "Pending / Action Required"
                            }
                            style={
                              isCurrentOpenSite
                                ? {
                                    backgroundColor: "#17324a",
                                    boxShadow: "inset 4px 0 0 #f0b429",
                                  }
                                : undefined
                            }
                          >
                            <td
                              className="billings-page__verification-table-select-cell"
                              style={
                                isCurrentOpenSite
                                  ? { backgroundColor: "#17324a" }
                                  : undefined
                              }
                            >
                              <input
                                type="checkbox"
                                checked={isSelected}
                                onChange={(event) =>
                                  toggleVerificationSelection(
                                    record.recordId,
                                    event.target.checked,
                                  )
                                }
                                aria-label={`Select ${record.screenCode}`}
                                title={
                                  isSelectable
                                    ? "Ready for Verify & Submit"
                                    : isPending
                                      ? "Open Billing Information & Invoice, complete the required site billing details, and click Validate first."
                                      : "First billing already completed."
                                }
                                disabled={!isSelectable}
                              />
                            </td>
                            <td
                              title={record.screenCode || "-"}
                              style={{
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                whiteSpace: "nowrap",
                                ...(isCurrentOpenSite
                                  ? { backgroundColor: "#17324a" }
                                  : {}),
                              }}
                            >
                              {record.screenCode || "-"}
                            </td>
                            <td
                              title={record.screenName || "-"}
                              style={{
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                whiteSpace: "nowrap",
                                ...(isCurrentOpenSite
                                  ? { backgroundColor: "#17324a" }
                                  : {}),
                              }}
                            >
                              {record.screenName || "-"}
                            </td>
                            <td
                              title={record.billingStartDate || "-"}
                              style={{
                                whiteSpace: "nowrap",
                                ...(isCurrentOpenSite
                                  ? { backgroundColor: "#17324a" }
                                  : {}),
                              }}
                            >
                              {record.billingStartDate || "-"}
                            </td>
                            <td
                              title={
                                getFirstTimeDisplaySubscriptionFee(
                                  record,
                                  billingRecords,
                                  stage3Records,
                                  commonComplexAllocations,
                                ) || "-"
                              }
                              style={{
                                whiteSpace: "nowrap",
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                ...(isCurrentOpenSite
                                  ? { backgroundColor: "#17324a" }
                                  : {}),
                              }}
                            >
                              {isVariableSubscriptionRecord(record)
                                ? "Variable"
                                : (() => {
                                    const fee =
                                      getFirstTimeDisplaySubscriptionFee(
                                        record,
                                        billingRecords,
                                        stage3Records,
                                        commonComplexAllocations,
                                      );
                                    return fee ? `₹${fee}` : "-";
                                  })()}
                            </td>
                            <td
                              style={
                                isCurrentOpenSite
                                  ? { backgroundColor: "#17324a" }
                                  : undefined
                              }
                            >
                              {isFirstBillingCompletedThisMonth
                                ? "Completed"
                                : isValidated
                                  ? "Validated"
                                  : "Pending"}
                            </td>
                            <td
                              className="billings-page__verification-table-action-cell"
                              style={
                                isCurrentOpenSite
                                  ? { backgroundColor: "#17324a" }
                                  : undefined
                              }
                            >
                              {isFirstBillingCompletedThisMonth ? (
                                <span
                                  title="First Time Billing is already processed. Later commercial changes must be initiated from Billing Records → Action → Price Change & Reallocation."
                                >
                                  -
                                </span>
                              ) : (
                                <select
                                  value=""
                                  onChange={(event) => {
                                    const action = event.target.value;

                                    if (!action) {
                                      return;
                                    }

                                    handleFirstTimeBillingRowAction(record, action);
                                    event.target.value = "";
                                  }}
                                  aria-label={`Action for ${record.screenCode}`}
                                  title="Select Action"
                                  style={{
                                    width: "100%",
                                    minWidth: 0,
                                    maxWidth: "100%",
                                  }}
                                >
                                  <option value="">Select</option>
                                  <option value="Open">Open</option>
                                  <option value="Stage 1">Move Stage 1</option>
                                  <option value="Stage 2">Move Stage 2</option>
                                </select>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    gap: 12,
                    marginTop: 12,
                    flexWrap: "wrap",
                  }}
                >
                  <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                    <span className="billings-page__helper" style={{ margin: 0 }}>
                      Showing {pagedVerificationRows.length} of {firstTimeVisibleRows.length} records
                    </span>
                    <div className="billings-page__field" style={{ marginBottom: 0, minWidth: 110 }}>
                                            <select
                        id="stage3-first-time-rpp-bottom"
                        value={firstTimeRowsPerPage}
                        onChange={(event) =>
                          setFirstTimeRowsPerPage(Number(event.target.value))
                        }
                      >
                        {rowsPerPageOptions.map((option) => (
                          <option key={option} value={option}>{option}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <button
                      type="button"
                      className="billings-page__secondary-button"
                      onClick={() => setFirstTimePage((page) => Math.max(1, page - 1))}
                      disabled={firstTimePage <= 1}
                    >
                      Previous
                    </button>
                    <span>Page {firstTimePage} of {firstTimePageCount}</span>
                    <button
                      type="button"
                      className="billings-page__secondary-button"
                      onClick={() =>
                        setFirstTimePage((page) => Math.min(firstTimePageCount, page + 1))
                      }
                      disabled={firstTimePage >= firstTimePageCount}
                    >
                      Next
                    </button>
                  </div>
                </div>
              </>
            )}
          </>
        )}
      </article>

      {activeBillingSection === billingSections.FIRST_TIME &&
        renderBillingInformationAndInvoice()}

      <article
        className="billings-page__card billings-page__verification-card"
        ref={priceChangeReallocationRef}
      >
        <div
          className="billings-page__card-header"
          style={{ display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%", gap: 16 }}
        >
          <div>
            <p className="billings-page__card-title">Price Change &amp; Reallocation</p>
            <span className="billings-page__card-badge billings-page__card-badge--editable">
              {billingSummaryLoading
                ? "Loading..."
                : `${priceChangeRows.length} ${priceChangeRows.length === 1 ? "Record" : "Records"}`}
            </span>
          </div>
          <button
            type="button"
            className="billings-page__secondary-button"
            onClick={() => handleBillingSectionToggle(billingSections.REALLOCATION)}
          >
            {activeBillingSection === billingSections.REALLOCATION ? "Minimize" : "Expand"}
          </button>
        </div>

        {activeBillingSection === billingSections.REALLOCATION && (
          <>
            <p className="billings-page__helper">
              Shows commercial price changes only after the site's Billing Start Date has been reached. Before first billing starts, corrections remain in First Time Billing. Common complex changes require a fresh site-wise reallocation; FoC screens are excluded.
            </p>

            <div className="billings-page__verification-actions" style={{ alignItems: "end", flexWrap: "wrap" }}>
              <div className="billings-page__field" style={{ flex: "1 1 420px", marginBottom: 0 }}>
                <label htmlFor="stage3-price-change-search">Search</label>
                <input
                  id="stage3-price-change-search"
                  type="search"
                  value={priceChangeSearchTerm}
                  onChange={(event) => setPriceChangeSearchTerm(event.target.value)}
                  placeholder="Search Billing Code, Complex Code, Screen Code, screen name or change type"
                />
              </div>
            </div>

            {priceChangeRows.length === 0 ? (
              <div className="billings-page__verification-empty">
                No price change or reallocation is currently required.
              </div>
            ) : (
              <div className="billings-page__verification-table-wrapper">
                <table className="billings-page__verification-table">
                  <thead>
                    <tr>
                      <SortableHeader label="Billing Code" sortKey="billingCode" activeKey={priceChangeSortKey} direction={priceChangeSortOrder} onSort={handlePriceChangeColumnSort} />
                      <SortableHeader label="Complex Code" sortKey="complexCode" activeKey={priceChangeSortKey} direction={priceChangeSortOrder} onSort={handlePriceChangeColumnSort} />
                      <SortableHeader label="Affected Screens" sortKey="screenCode" activeKey={priceChangeSortKey} direction={priceChangeSortOrder} onSort={handlePriceChangeColumnSort} />
                      <SortableHeader label="Change Type" sortKey="changeType" activeKey={priceChangeSortKey} direction={priceChangeSortOrder} onSort={handlePriceChangeColumnSort} />
                      <SortableHeader label="Previous Fee" sortKey="previousFee" activeKey={priceChangeSortKey} direction={priceChangeSortOrder} onSort={handlePriceChangeColumnSort} />
                      <SortableHeader label="Current Fee" sortKey="currentFee" activeKey={priceChangeSortKey} direction={priceChangeSortOrder} onSort={handlePriceChangeColumnSort} />
                      <SortableHeader label="Billing Start Date" sortKey="billingStartDate" activeKey={priceChangeSortKey} direction={priceChangeSortOrder} onSort={handlePriceChangeColumnSort} />
                      <SortableHeader label="Status" sortKey="status" activeKey={priceChangeSortKey} direction={priceChangeSortOrder} onSort={handlePriceChangeColumnSort} />
                      <SortableHeader label="Site Count" sortKey="affectedSites" activeKey={priceChangeSortKey} direction={priceChangeSortOrder} onSort={handlePriceChangeColumnSort} />
                      <th className="billings-page__foc-action-header">Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {priceChangeRows.map((row) => (
                      <tr key={row.rowKey}>
                        <td>{row.billingCode || "-"}</td>
                        <td>{getDisplayComplexCode(row.complexCode)}</td>
                        <td>{row.screenCode || "-"}</td>
                        <td>{row.changeType}</td>
                        <td>{row.previousFee ? `₹${row.previousFee}` : "-"}</td>
                        <td>{row.currentFee ? `₹${row.currentFee}` : "-"}</td>
                        <td>{formatBillingDateDisplay(row.billingStartDate)}</td>
                        <td>{row.status || "-"}</td>
                        <td>{row.affectedSites}</td>
                        <td>
                          <button
                            type="button"
                            className="billings-page__edit-button"
                            onClick={() => handleOpenPriceChange(row)}
                          >
                            Reallocate
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {activeRecord && activeBillingSection === billingSections.REALLOCATION ? (
              isCommonComplexBillingRecord(activeRecord) && commonComplexBillingSummary ? (
                <section className="billings-page__billing-section billings-page__billing-section--common-allocation" style={{ marginTop: 18 }}>
                  <div className="billings-page__billing-section-heading billings-page__billing-section-heading--stacked">
                    <div>
                      <h2>Reallocate Common Subscription Fee</h2>
                    </div>
                    <span className="billings-page__billing-section-text">
                      Previous allocations remain visible until the new allocations total the current Common Subscription Fee. Only non-FoC billable screens participate.
                    </span>
                  </div>

                  <p className="billings-page__modal-eyebrow billings-page__price-change-step-heading">
                    Step 1 - Price Change Details
                  </p>
                  <div className="billings-page__billing-allocation-summary">
                    <div className="billings-page__field"><label>Billing Code / Customer Code</label><input type="text" value={commonComplexBillingSummary.billingCode} readOnly /></div>
                    <div className="billings-page__field"><label>Complex Code</label><input type="text" value={getDisplayComplexCode(commonComplexBillingSummary.complexCode)} readOnly /></div>
                    <div className="billings-page__field"><label>Current Common Fee</label><input type="text" value={commonComplexBillingSummary.commonSubscriptionFee} readOnly /></div>
                    <div className="billings-page__field"><label>Currently Billable Sites</label><input type="text" value={commonComplexBillingSummary.currentlyBillableSites} readOnly /></div>
                    <div className="billings-page__field"><label>Allocation Total</label><input type="text" value={commonComplexBillingSummary.allocationTotal} readOnly /></div>
                    <div className="billings-page__field"><label>Difference</label><input type="text" value={commonComplexBillingSummary.difference} readOnly /></div>
                    <div className="billings-page__field"><label>Change Type *</label><select value={priceChangeDraft.changeType} onChange={(event) => handlePriceChangeTypeSelection(event.target.value)}><option value="Price Change">Price Change</option><option value="Mode Change">Mode Change</option>{normalizeSubscriptionModeLabel(activeRecord.subscriptionMode) !== "Monthly" && (<option value="Price + Mode Change">Price + Mode Change</option>)}</select></div>
                    <div className="billings-page__field"><label>Current Subscription Mode</label><input type="text" value={activeRecord.subscriptionMode || "Monthly"} readOnly /></div>
                    {(priceChangeDraft.changeType === "Mode Change" || priceChangeDraft.changeType === "Price + Mode Change") && (
                      <div className="billings-page__field"><label>New Subscription Mode *</label><select value={priceChangeDraft.newMode} onChange={(event) => handleNewSubscriptionModeSelection(event.target.value)}><option value="Monthly">Monthly</option><option value="Half-Yearly">Half-Yearly</option><option value="Annual">Annual</option></select></div>
                    )}
                    <div className="billings-page__field"><label>New Common Subscription Fee *</label><input type="text" inputMode="decimal" value={priceChangeDraft.newFee} disabled={Boolean(getPriceChangeInvoiceLock(activeRecord, parseLocalDateValue(priceChangeDraft.effectiveDate)))} onChange={(event) => { setPriceChangeMessage(""); setPriceChangeDraft((current) => ({ ...current, newFee: sanitizeAmount(event.target.value) })); }} onBlur={handleNewSubscriptionFeeBlur} /></div>
                    <div className="billings-page__field"><label>Price Change Effective Date *</label><input
                      type="text"
                      inputMode="numeric"
                      maxLength={10}
                      placeholder="DD-MM-YYYY"
                      value={priceChangeEffectiveDateDisplay}
                      onChange={handlePriceChangeEffectiveDateInput}
                      onBlur={handlePriceChangeEffectiveDateBlur}
                      aria-label="Price Change Effective Date"
                    /></div>
                    {getPriceChangeInvoiceLock(activeRecord, parseLocalDateValue(priceChangeDraft.effectiveDate)) && (
                      <p className="billings-page__helper" style={{ gridColumn: "1 / -1" }}>
                        Price and Effective Date are locked by the latest ERP invoice state for this billing period. Only ERP-confirmed revoked/cancelled state with no valid replacement unlocks correction.
                      </p>
                    )}
                    <div className="billings-page__field" style={{ gridColumn: "1 / -1" }}><label>Remarks / Reason *</label><textarea rows={3} value={priceChangeDraft.remarks} onChange={(event) => setPriceChangeDraft((current) => ({ ...current, remarks: event.target.value }))} placeholder="Enter reason for the price change / reallocation" /></div>
                  </div>


                  {activePriceChangeAdjustmentPreview && (
                    <div className="billings-page__billing-allocation-summary" style={{ marginTop: 14 }}>
                      {activePriceChangeAdjustmentPreview.commonGroupAdjustment && (
                        <>
                          <div className="billings-page__field"><label>Previous Common Fee</label><input type="text" value={`₹${activePriceChangeAdjustmentPreview.previousCommonFee}`} readOnly /></div>
                          <div className="billings-page__field"><label>Utilized Amount up to Effective Date - 1</label><input type="text" value={`₹${activePriceChangeAdjustmentPreview.utilizedOldPeriod}`} readOnly /></div>
                          <div className="billings-page__field"><label>Unused / Balance Previous Entitlement</label><input type="text" value={`₹${activePriceChangeAdjustmentPreview.unusedOldEntitlement}`} readOnly /></div>
                        </>
                      )}
                      <div className="billings-page__field"><label>Current / Existing Billing Cycle</label><input type="text" value={`${formatBillingDateDisplay(activePriceChangeAdjustmentPreview.cycleFrom)} to ${formatBillingDateDisplay(activePriceChangeAdjustmentPreview.cycleTo)}`} readOnly /></div>
                      {!activePriceChangeAdjustmentPreview.commonGroupAdjustment && (
                        <>
                          <div className="billings-page__field"><label>Amount Already Billed</label><input type="text" value={activePriceChangeAdjustmentPreview.amountAlreadyBilled ? `₹${activePriceChangeAdjustmentPreview.amountAlreadyBilled}` : "-"} readOnly /></div>
                          <div className="billings-page__field"><label>{activePriceChangeAdjustmentPreview.modeChange ? "Old Mode Entitlement up to Effective Date - 1" : "Revised Pro-rata Entitlement"}</label><input type="text" value={activePriceChangeAdjustmentPreview.revisedEntitlement ? `₹${activePriceChangeAdjustmentPreview.revisedEntitlement}` : "-"} readOnly /></div>
                        </>
                      )}
                      <div className="billings-page__field"><label>Adjustment Type</label><input type="text" value={activePriceChangeAdjustmentPreview.adjustmentType} readOnly /></div>
                      <div className="billings-page__field"><label>Credit / Balance Amount</label><input type="text" value={activePriceChangeAdjustmentPreview.adjustmentAmount ? `₹${activePriceChangeAdjustmentPreview.adjustmentAmount}` : "₹0"} readOnly /></div>
                      {activePriceChangeAdjustmentPreview.modeChange && (
                        <>
                          <div className="billings-page__field"><label>Old Mode End Date</label><input type="text" value={formatBillingDateDisplay(activePriceChangeAdjustmentPreview.oldModeEnd)} readOnly /></div>
                          <div className="billings-page__field"><label>New Mode First Cycle</label><input type="text" value={`${formatBillingDateDisplay(activePriceChangeAdjustmentPreview.newCycleFrom)} to ${formatBillingDateDisplay(activePriceChangeAdjustmentPreview.newCycleTo)}`} readOnly /></div>
                          <div className="billings-page__field"><label>New Mode First Cycle Charge</label><input type="text" value={activePriceChangeAdjustmentPreview.newCycleFee ? `₹${activePriceChangeAdjustmentPreview.newCycleFee}` : "-"} readOnly /></div>
                          <div className="billings-page__field"><label>Net Next Billing After Credit/Balance</label><input type="text" value={activePriceChangeAdjustmentPreview.netNextBilling ? `₹${activePriceChangeAdjustmentPreview.netNextBilling}` : "-"} readOnly /></div>
                        </>
                      )}
                    </div>
                  )}

                  <p className="billings-page__modal-eyebrow billings-page__price-change-step-heading">
                    Step 2 - Site-wise Allocation
                  </p>
                  <div
                    className="billings-page__billing-allocation-table-wrapper"
                    ref={priceChangeAllocationRef}
                  >
                    <table className="billings-page__billing-allocation-table">
                      <thead>
                        <tr>
                          <th>Screen Code</th>
                          <th>Screen / Complex Name</th>
                          <th>Previous / Current Allocation</th>
                          <th>New Allocation</th>
                          <th>Current Billing Charge</th>
                        </tr>
                      </thead>
                      <tbody>
                        {commonComplexBillingRows.map((row) => (
                          <tr key={`reallocation-${row.siteKey}`}>
                            <td>{row.screenCode || "-"}</td>
                            <td>{row.screenName || "-"}</td>
                            <td>{reallocationOriginalAllocations[row.siteKey] ? `₹${reallocationOriginalAllocations[row.siteKey]}` : "-"}</td>
                            <td>
                              <input
                                type="text"
                                inputMode="decimal"
                                value={row.allocatedFee || ""}
                                onChange={(event) => handleCommonComplexAllocationInlineChange(row.siteKey, event.target.value)}
                                readOnly={false}
                                placeholder="Enter new allocation"
                                style={{ width: "100%", minWidth: 0, boxSizing: "border-box" }}
                              />
                            </td>
                            <td>{row.currentBillingCharge || "-"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <div className="billings-page__detail-actions billings-page__price-change-actions">
                    <button
                      type="button"
                      className="billings-page__secondary-button"
                      onClick={handleCancelPriceChange}
                    >
                      Cancel
                    </button>
                    <button type="button" className="billings-page__primary-button" onClick={handleValidatePriceChange} disabled={isApplyingPriceChange}>
                      {isApplyingPriceChange ? "Applying Price Change..." : "Apply Price Change"}
                    </button>
                  </div>
                </section>
              ) : isAllocationModelBillingRecord(activeRecord) ? (
                <section className="billings-page__billing-section" style={{ marginTop: 18 }}>
                  <div className="billings-page__billing-section-heading">
                    <h2>Allocation Model Change</h2>
                    <span className="billings-page__billing-section-text">
                      The Current Allocation is shown read-only. Select the New Allocation manually, then select the New Plan manually for each eligible screen. Validate the completed change first; Save & Submit is the final action and becomes available only for the exact validated values. The New Subscription Fee is fetched from the selected Plan. OTF is not changed in this process.
                    </span>
                  </div>

                  <div className="billings-page__billing-section-grid">
                    <div className="billings-page__field"><label>Billing Code</label><input type="text" value={activeRecord.billingCode || ""} readOnly /></div>
                    <div className="billings-page__field"><label>Chain Network</label><input type="text" value={getAllocationChainName(activeRecord) || "-"} readOnly /></div>
                    <div className="billings-page__field"><label>Pricing Method</label><input type="text" value="Allocation Model" readOnly /></div>
                    <div className="billings-page__field"><label>Affected Sites / Complexes</label><input type="text" value={getAllocationModelAffectedRecords(activeRecord).length} readOnly /></div>
                    <div className="billings-page__field" style={{ gridColumn: "1 / -1" }}><label>Current Allocation Model</label><input type="text" value={formatAllocationModelSummary(getAllocationModelCurrentSnapshot(activeRecord))} readOnly /></div>
                    <div className="billings-page__field"><label>New Allocation *</label>
                      <select
                        value={allocationModelChangeDraft.modelId}
                        onChange={(event) => {
                          setPriceChangeMessage("");
                          setAllocationModelChangeDraft((current) => ({
                            ...current,
                            modelId: event.target.value,
                            planSelections: {},
                          }));
                        }}
                      >
                        <option value="">Select Active Chain Allocation</option>
                        {getConfiguredAllocationModelsForRecord(activeRecord, configuredAllocationModels).map((model) => (
                          <option key={model.modelId} value={model.modelId}>
                            {getAllocationDisplayName(model)}{model.effectiveFrom ? ` · ${formatBillingDateDisplay(model.effectiveFrom)}` : ""}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div className="billings-page__field"><label>Effective Date *</label><input type="date" value={allocationModelChangeDraft.effectiveDate} onChange={(event) => { setPriceChangeMessage(""); setAllocationModelChangeDraft((current) => ({ ...current, effectiveDate: event.target.value })); }} /></div>
                    <div className="billings-page__field" style={{ gridColumn: "1 / -1" }}><label>New Model Plan Fee Structure</label><input type="text" value={getSelectedAllocationModel() ? formatAllocationModelSummary(getSelectedAllocationModel()) : ""} readOnly placeholder="Select New Allocation first" /></div>
                    <div className="billings-page__field" style={{ gridColumn: "1 / -1" }}><label>Remarks / Reason *</label><textarea rows={3} value={allocationModelChangeDraft.remarks} onChange={(event) => setAllocationModelChangeDraft((current) => ({ ...current, remarks: event.target.value }))} placeholder="Enter reason for the Allocation Model change" /></div>
                  </div>

                  <div className="billings-page__billing-allocation-table-wrapper" style={{ marginTop: 14 }}>
                    <table className="billings-page__billing-allocation-table">
                      <thead>
                        <tr>
                          <th>Screen Code</th>
                          <th>Screen / Complex Name</th>
                          <th>Device Count</th>
                          <th>Current Plan</th>
                          <th>Current Fee</th>
                          <th>New Plan</th>
                          <th>New Fee</th>
                        </tr>
                      </thead>
                      <tbody>
                        {getAllocationModelAffectedRecords(activeRecord).map((record) => {
                          const currentRow = getAllocationModelRowForRecord(record);
                          const deviceCount =
                            Number(currentRow?.count) ||
                            Number(record?.deviceCount) ||
                            1;
                          const selectedModel = getSelectedAllocationModel();
                          const selectedPlanLabel = normalizeValue(
                            allocationModelChangeDraft.planSelections?.[record.recordId],
                          );
                          const selectedPlan =
                            selectedModel?.slabs?.find(
                              (slab) =>
                                normalizeValue(slab?.label) === selectedPlanLabel,
                            ) || null;

                          return (
                            <tr key={`allocation-model-change-${record.recordId}`}>
                              <td>{record.screenCode || "-"}</td>
                              <td>{record.screenName || record.billingName || "-"}</td>
                              <td>{deviceCount}</td>
                              <td>{currentRow?.planMode || "-"}</td>
                              <td>{currentRow?.planFee ? `₹${currentRow.planFee}` : "-"}</td>
                              <td>
                                <select
                                  value={selectedPlanLabel}
                                  disabled={!selectedModel}
                                  onChange={(event) => {
                                    setPriceChangeMessage("");
                                    const nextPlan = event.target.value;
                                    setAllocationModelChangeDraft((current) => ({
                                      ...current,
                                      planSelections: {
                                        ...(current.planSelections || {}),
                                        [record.recordId]: nextPlan,
                                      },
                                    }));
                                  }}
                                >
                                  <option value="">
                                    {selectedModel ? "Select New Plan" : "Select New Allocation first"}
                                  </option>
                                  {(selectedModel?.slabs || []).map((slab) => (
                                    <option key={slab.slabId} value={slab.label}>
                                      {slab.label}
                                    </option>
                                  ))}
                                </select>
                              </td>
                              <td>
                                {selectedPlan?.planFee
                                  ? `₹${selectedPlan.planFee}`
                                  : "-"}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>

                  {getConfiguredAllocationModelsForRecord(activeRecord, configuredAllocationModels).length === 0 && (
                    <p className="billings-page__helper" style={{ marginTop: 12 }}>
                      No Active Allocation is configured in Settings for {getAllocationChainName(activeRecord) || "this Chain Network"}.
                    </p>
                  )}

                  <div className="billings-page__detail-actions" style={{ justifyContent: "flex-end", marginTop: 14, gap: 10 }}>
                    <button
                      type="button"
                      className="billings-page__secondary-button"
                      onClick={handleCancelPriceChange}
                    >
                      Cancel
                    </button>
                    <button
                      type="button"
                      className="billings-page__secondary-button"
                      onClick={handleValidateAllocationModelChange}
                    >
                      Validate &amp; Apply Allocation Model
                    </button>
                    <button
                      type="button"
                      className="billings-page__primary-button"
                      onClick={handleSaveAndSubmitAllocationModelChange}
                      disabled={!isAllocationModelValidationCurrent()}
                      title={
                        isAllocationModelValidationCurrent()
                          ? "Save and submit the validated Allocation Model change"
                          : "Validate the current Allocation Model values before Save & Submit"
                      }
                    >
                      Save &amp; Submit
                    </button>
                  </div>
                </section>
              ) : (
                <section className="billings-page__billing-section" style={{ marginTop: 18 }}>
                  <div className="billings-page__billing-section-heading">
                    <h2>Price Change Review</h2>
                    <span className="billings-page__billing-section-text">
                      This standalone or Site / Screen-wise price change does not require complex allocation. The revised Subscription Fee is handled as the current commercial value while previous billing history remains unchanged.
                    </span>
                  </div>
                  <div className="billings-page__billing-section-grid">
                    <div className="billings-page__field"><label>Screen Code</label><input type="text" value={activeRecord.screenCode || ""} readOnly /></div>
                    <div className="billings-page__field"><label>Change Type *</label><select value={priceChangeDraft.changeType} onChange={(event) => handlePriceChangeTypeSelection(event.target.value)}><option value="Price Change">Price Change</option><option value="Mode Change">Mode Change</option>{normalizeSubscriptionModeLabel(activeRecord.subscriptionMode) !== "Monthly" && (<option value="Price + Mode Change">Price + Mode Change</option>)}</select></div>
                    <div className="billings-page__field"><label>Current Subscription Mode</label><input type="text" value={activeRecord.subscriptionMode || "Monthly"} readOnly /></div>
                    {(priceChangeDraft.changeType === "Mode Change" || priceChangeDraft.changeType === "Price + Mode Change") && (
                      <div className="billings-page__field"><label>New Subscription Mode *</label><select value={priceChangeDraft.newMode} onChange={(event) => handleNewSubscriptionModeSelection(event.target.value)}><option value="Monthly">Monthly</option><option value="Half-Yearly">Half-Yearly</option><option value="Annual">Annual</option></select></div>
                    )}
                    <div className="billings-page__field"><label>Current Subscription Fee</label><input type="text" value={resolveSiteWiseSubscriptionFee(activeRecord) || activeRecord.subscriptionFee || ""} readOnly /></div>
                    <div className="billings-page__field"><label>New Subscription Fee *</label><input type="text" inputMode="decimal" value={priceChangeDraft.newFee} onChange={(event) => { setPriceChangeMessage(""); setPriceChangeDraft((current) => ({ ...current, newFee: sanitizeAmount(event.target.value) })); }} onBlur={handleNewSubscriptionFeeBlur} /></div>
                    <div className="billings-page__field"><label>Price Change Effective Date *</label><input
                      type="text"
                      inputMode="numeric"
                      maxLength={10}
                      placeholder="DD-MM-YYYY"
                      value={priceChangeEffectiveDateDisplay}
                      onChange={handlePriceChangeEffectiveDateInput}
                      onBlur={handlePriceChangeEffectiveDateBlur}
                      aria-label="Price Change Effective Date"
                    /></div>
                    <div className="billings-page__field" style={{ gridColumn: "1 / -1" }}><label>Remarks / Reason *</label><textarea rows={3} value={priceChangeDraft.remarks} onChange={(event) => setPriceChangeDraft((current) => ({ ...current, remarks: event.target.value }))} placeholder="Enter reason for the price change" /></div>
                  </div>
                  {activePriceChangeAdjustmentPreview && (
                    <div className="billings-page__billing-allocation-summary" style={{ marginTop: 14 }}>
                      <div className="billings-page__field"><label>Current / Existing Billing Cycle</label><input type="text" value={`${formatBillingDateDisplay(activePriceChangeAdjustmentPreview.cycleFrom)} to ${formatBillingDateDisplay(activePriceChangeAdjustmentPreview.cycleTo)}`} readOnly /></div>
                      <div className="billings-page__field"><label>Amount Already Billed</label><input type="text" value={activePriceChangeAdjustmentPreview.amountAlreadyBilled ? `₹${activePriceChangeAdjustmentPreview.amountAlreadyBilled}` : "-"} readOnly /></div>
                      <div className="billings-page__field"><label>{activePriceChangeAdjustmentPreview.modeChange ? "Old Mode Entitlement up to Effective Date - 1" : "Revised Pro-rata Entitlement"}</label><input type="text" value={activePriceChangeAdjustmentPreview.revisedEntitlement ? `₹${activePriceChangeAdjustmentPreview.revisedEntitlement}` : "-"} readOnly /></div>
                      <div className="billings-page__field"><label>Adjustment Type</label><input type="text" value={activePriceChangeAdjustmentPreview.adjustmentType} readOnly /></div>
                      <div className="billings-page__field"><label>Credit / Balance Amount</label><input type="text" value={activePriceChangeAdjustmentPreview.adjustmentAmount ? `₹${activePriceChangeAdjustmentPreview.adjustmentAmount}` : "₹0"} readOnly /></div>
                      {activePriceChangeAdjustmentPreview.modeChange && (
                        <>
                          <div className="billings-page__field"><label>Old Mode End Date</label><input type="text" value={formatBillingDateDisplay(activePriceChangeAdjustmentPreview.oldModeEnd)} readOnly /></div>
                          <div className="billings-page__field"><label>New Mode First Cycle</label><input type="text" value={`${formatBillingDateDisplay(activePriceChangeAdjustmentPreview.newCycleFrom)} to ${formatBillingDateDisplay(activePriceChangeAdjustmentPreview.newCycleTo)}`} readOnly /></div>
                          <div className="billings-page__field"><label>New Mode First Cycle Charge</label><input type="text" value={activePriceChangeAdjustmentPreview.newCycleFee ? `₹${activePriceChangeAdjustmentPreview.newCycleFee}` : "-"} readOnly /></div>
                          <div className="billings-page__field"><label>Net Next Billing After Credit/Balance</label><input type="text" value={activePriceChangeAdjustmentPreview.netNextBilling ? `₹${activePriceChangeAdjustmentPreview.netNextBilling}` : "-"} readOnly /></div>
                        </>
                      )}
                    </div>
                  )}
                  <div className="billings-page__detail-actions billings-page__price-change-actions">
                    <button
                      type="button"
                      className="billings-page__secondary-button"
                      onClick={handleCancelPriceChange}
                    >
                      Cancel
                    </button>
                    <button type="button" className="billings-page__primary-button" onClick={handleValidatePriceChange} disabled={isApplyingPriceChange}>
                      {isApplyingPriceChange ? "Applying Price Change..." : "Apply Price Change"}
                    </button>
                  </div>
                </section>
              )
            ) : null}
            {activeBillingSection === billingSections.REALLOCATION && priceChangeMessage && (
              <p className="billings-page__helper" style={{ marginTop: 12, fontWeight: 600 }}>
                {priceChangeMessage}
              </p>
            )}
          </>
        )}
      </article>

      <article className="billings-page__card billings-page__verification-card">
        <div className="billings-page__card-header" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%", gap: 16 }}>
          <div>
            <p className="billings-page__card-title">Recurring Billing</p>
            <span className="billings-page__card-badge billings-page__card-badge--editable">
              {billingSummaryLoading
                ? "Loading..."
                : `${formatSummaryCount(canonicalBillingSummary?.recurringBilling)} Records`}
            </span>
          </div>
          <div style={{ display: "flex", alignItems: "end", gap: 10, marginLeft: "auto", flexShrink: 0 }}>
            <button
              type="button"
              className="billings-page__secondary-button"
              onClick={() => handleBillingSectionToggle(billingSections.RECURRING)}
            >
              {activeBillingSection === billingSections.RECURRING ? "Minimize" : "Expand"}
            </button>
          </div>
        </div>

        {activeBillingSection === billingSections.RECURRING && (
          <>
            <p className="billings-page__helper">
              Recurring Billing shows the oldest unbilled/overdue cycle first, followed by the current due cycle. Future cycles remain hidden. Price Change &amp; Reallocation and Pause/Inactive controls are respected before processing.
            </p>

            <div className="billings-page__verification-actions" style={{ alignItems: "end", flexWrap: "wrap" }}>
              <div className="billings-page__field" style={{ flex: "1 1 420px", marginBottom: 0 }}>
                <label htmlFor="stage3-recurring-search">Search</label>
                <input
                  id="stage3-recurring-search"
                  type="search"
                  value={recurringSearchTerm}
                  onChange={(event) => setRecurringSearchTerm(event.target.value)}
                  placeholder="Search code/name/location or type unbilled, pending, overdue, billed, processed"
                />
              </div>

              <button
                type="button"
                className="billings-page__secondary-button"
                onClick={handleDownloadRecurringBillingExcel}
                disabled={recurringBillingDownloadRows.length === 0}
              >
                Download Excel
              </button>
            </div>

            {filteredRecurringRows.length === 0 ? (
              <div className="billings-page__verification-empty">
                No recurring billing sites are available yet.
              </div>
            ) : (
              <>
                <div className="billings-page__verification-table-wrapper">
                  <table className="billings-page__verification-table">
                    <colgroup>
                      <col style={{ width: "112px" }} />
                      <col />
                      <col />
                      <col style={{ width: "112px" }} />
                      <col style={{ width: "140px" }} />
                      <col style={{ width: "84px" }} />
                    </colgroup>
                    <thead>
                      <tr>
                        <SortableHeader label="Screen Code" sortKey="screenCode" activeKey={recurringSortKey} direction={recurringSortOrder} onSort={(key) => toggleColumnSort(recurringSortKey, setRecurringSortKey, recurringSortOrder, setRecurringSortOrder, key)} />
                        <SortableHeader label="Screen Name" sortKey="screenName" activeKey={recurringSortKey} direction={recurringSortOrder} onSort={(key) => toggleColumnSort(recurringSortKey, setRecurringSortKey, recurringSortOrder, setRecurringSortOrder, key)} />
                        <SortableHeader label="Location" sortKey="location" activeKey={recurringSortKey} direction={recurringSortOrder} onSort={(key) => toggleColumnSort(recurringSortKey, setRecurringSortKey, recurringSortOrder, setRecurringSortOrder, key)} />
                        <SortableHeader label="Billing Date" sortKey="billingDate" activeKey={recurringSortKey} direction={recurringSortOrder} onSort={(key) => toggleColumnSort(recurringSortKey, setRecurringSortKey, recurringSortOrder, setRecurringSortOrder, key)} />
                        <th>Status</th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {pagedRecurringRows.map((record) => (
                        <tr key={record.recordId}>
                          <td title={record.screenCode || "-"}>{record.screenCode || "-"}</td>
                          <td title={record.screenName || "-"}>{record.screenName || "-"}</td>
                          <td title={record.location || "-"}>{record.location || "-"}</td>
                          <td title={record.recurringCycle?.periodFrom || "-"}>
                            {record.recurringCycle?.periodFrom || "-"}
                          </td>
                          <td title={record.recurringCycle?.eligibilityStatus || "-"}>
                            {record.recurringCycle?.eligibilityStatus || "-"}
                          </td>
                          <td>
                            <button
                              type="button"
                              className="billings-page__verification-open-button"
                              onClick={() => handleOpenRecurringRecord(record.recordId)}
                            >
                              Open
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, marginTop: 12, flexWrap: "wrap" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                    <span className="billings-page__helper" style={{ margin: 0 }}>
                      Showing {pagedRecurringRows.length} of {filteredRecurringRows.length} records
                    </span>
                    <div className="billings-page__field" style={{ marginBottom: 0, minWidth: 110 }}>
                                            <select
                        id="stage3-recurring-rpp-bottom"
                        value={recurringRowsPerPage}
                        onChange={(event) => setRecurringRowsPerPage(Number(event.target.value))}
                      >
                        {rowsPerPageOptions.map((option) => (
                          <option key={option} value={option}>{option}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <button
                      type="button"
                      className="billings-page__secondary-button"
                      onClick={() => setRecurringPage((page) => Math.max(1, page - 1))}
                      disabled={recurringPage <= 1}
                    >
                      Previous
                    </button>
                    <span>Page {recurringPage} of {recurringPageCount}</span>
                    <button
                      type="button"
                      className="billings-page__secondary-button"
                      onClick={() => setRecurringPage((page) => Math.min(recurringPageCount, page + 1))}
                      disabled={recurringPage >= recurringPageCount}
                    >
                      Next
                    </button>
                  </div>
                </div>

                {activeRecurringRecord && (
                  <section
                    className="billings-page__billing-section"
                    ref={recurringBillingWorkspaceRef}
                    style={{ marginTop: 24 }}
                  >
                    <div className="billings-page__billing-section-heading">
                      <h2>3B Recurring Billing Details</h2>
                      <span className="billings-page__billing-section-text">
                        Process or hold the live recurring billing month. Fixed billing values are carried forward from approved billing history. Variable billing requires the current-month amount.
                      </span>
                    </div>

                    <div className="billings-page__billing-section-grid">
                      <div className="billings-page__field">
                        <label>Billing Code / Customer Code</label>
                        <input type="text" value={activeRecurringRecord.billingCode || ""} readOnly />
                      </div>

                      <div className="billings-page__field">
                        <label>Screen Code</label>
                        <input type="text" value={activeRecurringRecord.screenCode || ""} readOnly />
                      </div>

                      <div className="billings-page__field">
                        <label>Screen Name</label>
                        <input type="text" value={activeRecurringRecord.screenName || ""} readOnly />
                      </div>

                      <div className="billings-page__field">
                        <label>Location</label>
                        <input type="text" value={activeRecurringRecord.location || ""} readOnly />
                      </div>

                      <div className="billings-page__field">
                        <label>Billing Month</label>
                        <input type="text" value={activeRecurringRecord.recurringCycle?.billingMonth || ""} readOnly />
                      </div>

                      <div className="billings-page__field">
                        <label>Billing Period</label>
                        <input
                          type="text"
                          value={formatBillingPeriod(
                            activeRecurringRecord.recurringCycle?.periodFrom,
                            activeRecurringRecord.recurringCycle?.periodTo,
                          )}
                          readOnly
                        />
                      </div>

                      <div className="billings-page__field">
                        <label>Billing Status</label>
                        <input type="text" value={activeRecurringRecord.recurringCycle?.eligibilityStatus || ""} readOnly />
                      </div>

                      <div className="billings-page__field">
                        <label>Subscription Type</label>
                        <input type="text" value={activeRecurringRecord.subscriptionType || ""} readOnly />
                      </div>

                      <div className="billings-page__field">
                        <label>Subscription Mode</label>
                        <input type="text" value={activeRecurringRecord.subscriptionMode || ""} readOnly />
                      </div>

                      <div className="billings-page__field">
                        <label>Actual Subscription Fee</label>
                        <input
                          type="text"
                          value={activeRecurringRecord.recurringCycle?.actualSubscriptionFee || ""}
                          readOnly
                        />
                      </div>

                      <div className="billings-page__field">
                        <label>Billing Subscription Fee</label>
                        {normalizeValue(activeRecurringRecord.subscriptionType).toLowerCase() === "variable" ? (
                          <input
                            type="text"
                            inputMode="decimal"
                            value={
                              activeRecurringRecord?.monthlyBillingAmounts?.[
                                activeRecurringRecord.recurringCycle?.monthKey
                              ] || ""
                            }
                            onChange={(event) =>
                              handleVariableMonthlyAmountChange(
                                activeRecurringRecord.recurringCycle?.monthKey,
                                event.target.value,
                              )
                            }
                            placeholder="Enter current month amount"
                          />
                        ) : (
                          <input
                            type="text"
                            value={activeRecurringRecord.recurringCycle?.billingSubscriptionFee || ""}
                            readOnly
                          />
                        )}
                      </div>
                    </div>

                    <div className="billings-page__billing-section-heading" style={{ marginTop: 18 }}>
                      <h2>Billing Action</h2>
                    </div>

                    <div style={{ display: "flex", gap: 18, alignItems: "center", flexWrap: "wrap", marginBottom: 16 }}>
                      <label style={{ display: "flex", gap: 8, alignItems: "center" }}>
                        <input
                          type="radio"
                          name="recurring-billing-action"
                          value="Process"
                          checked={recurringDecision === "Process"}
                          onChange={() => setRecurringDecision("Process")}
                        />
                        Process
                      </label>

                      <label style={{ display: "flex", gap: 8, alignItems: "center" }}>
                        <input
                          type="radio"
                          name="recurring-billing-action"
                          value="Hold"
                          checked={recurringDecision === "Hold"}
                          onChange={() => setRecurringDecision("Hold")}
                        />
                        Hold
                      </label>
                    </div>

                    {recurringDecision === "Hold" && (
                      <div className="billings-page__field" style={{ maxWidth: 760 }}>
                        <label>Remarks *</label>
                        <textarea
                          value={recurringRemarks}
                          onChange={(event) => setRecurringRemarks(event.target.value)}
                          placeholder="Enter reason for holding this billing month"
                          rows={3}
                        />
                      </div>
                    )}

                    <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 18 }}>
                      <button
                        type="button"
                        className="billings-page__primary-button"
                        onClick={handleRecurringDecisionSubmit}
                      >
                        {recurringDecision === "Hold"
                          ? "Save Hold"
                          : "Process Recurring Billing"}
                      </button>
                    </div>
                  </section>
                )}
              </>
            )}
          </>
        )}
      </article>


      <article className="billings-page__card billings-page__verification-card billings-page__billing-section--billing-records">
        <div className="billings-page__card-header" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%", gap: 16 }}>
          <div>
            <p className="billings-page__card-title">Billing Records</p>
            <span className="billings-page__card-badge billings-page__card-badge--editable">
              {billingSummaryLoading
                ? "Loading..."
                : `${formatSummaryCount(canonicalBillingSummary?.billingRecords)} Records`}
            </span>
          </div>
          <div style={{ display: "flex", alignItems: "end", gap: 10, marginLeft: "auto", flexShrink: 0 }}>
            <button
              type="button"
              className="billings-page__secondary-button"
              onClick={() => handleBillingSectionToggle(billingSections.RECORDS)}
            >
              {activeBillingSection === billingSections.RECORDS ? "Minimize" : "Expand"}
            </button>
          </div>
        </div>

        {activeBillingSection === billingSections.RECORDS && (
          <>
            <p className="billings-page__helper">
              Site-wise billing records remain as permanent financial history. Status shows the site's current billing state; unresolved Inactive records are highlighted after 6 hours.
            </p>

            {billingRecordsCommonPricingSummaries.map((summary) => {
              const activeBillableLabel =
                summary.activeBillableScreens === 1
                  ? "Active Billable Screen"
                  : "Active Billable Screens";
              const currentBillingLabel =
                summary.activeBillableScreens > 1
                  ? "Current Combined Billing Amount"
                  : "Current Billing Amount";

              return (
                <section
                key={`common-pricing-summary-${summary.groupKey}`}
                aria-label={`Common Pricing Summary for ${
                  summary.complexCode || summary.billingCode
                }`}
                style={{
                  marginBottom: 14,
                  padding: "12px 14px",
                  border: "1px solid rgba(148, 163, 184, 0.24)",
                  borderRadius: 12,
                  background: "rgba(148, 163, 184, 0.06)",
                }}
              >
                <div
                  style={{
                    display: "flex",
                    alignItems: "baseline",
                    justifyContent: "space-between",
                    gap: 12,
                    flexWrap: "wrap",
                    marginBottom: 10,
                  }}
                >
                  <strong>Common Pricing Summary</strong>
                  <span className="billings-page__helper" style={{ margin: 0 }}>
                    {getDisplayComplexCode(summary.complexCode) ||
                      summary.billingCode}
                  </span>
                </div>
                <div
                  className="billings-page__common-pricing-summary-grid"
                  style={{
                    width: "100%",
                  }}
                >
                  {[
                    [activeBillableLabel, summary.activeBillableScreens],
                    [
                      currentBillingLabel,
                      formatInrAmount(summary.currentCombinedBillingAmount),
                    ],
                    [
                      "Configured Common Total",
                      formatInrAmount(summary.configuredCommonTotal),
                    ],
                    [
                      "Screens Awaiting Billing",
                      summary.screensAwaitingBilling,
                    ],
                    ["FoC Screens", summary.focScreens],
                  ].map(([label, value]) => (
                    <div
                      key={label}
                      className="billings-page__common-pricing-summary-item"
                    >
                      <span
                        className="billings-page__helper billings-page__common-pricing-summary-label"
                      >
                        {label}
                      </span>
                      <strong className="billings-page__common-pricing-summary-value">
                        {value}
                      </strong>
                    </div>
                  ))}
                </div>
                </section>
              );
            })}

            <div className="billings-page__verification-actions" style={{ alignItems: "end", flexWrap: "wrap", marginBottom: 16 }}>
              <div className="billings-page__field" style={{ flex: "1 1 420px", marginBottom: 0 }}>
                <label htmlFor="stage3-billing-records-search">Search</label>
                <input
                  id="stage3-billing-records-search"
                  type="search"
                  value={billingRecordsSearchTerm}
                  onChange={(event) => setBillingRecordsSearchTerm(event.target.value)}
                  placeholder="Search Customer Code, Complex Code, Screen Code, Screen Name, Location, or Status"
                />
              </div>

              <div className="billings-page__field" style={{ minWidth: 145, marginBottom: 0 }}>
                <label htmlFor="billing-records-from-date">From Date</label>
                <input id="billing-records-from-date" type="date" value={billingRecordsFromDate} onChange={(event) => setBillingRecordsFromDate(event.target.value)} />
              </div>

              <div className="billings-page__field" style={{ minWidth: 145, marginBottom: 0 }}>
                <label htmlFor="billing-records-to-date">To Date</label>
                <input id="billing-records-to-date" type="date" value={billingRecordsToDate} onChange={(event) => setBillingRecordsToDate(event.target.value)} />
              </div>

              <button
                type="button"
                className="billings-page__secondary-button"
                onClick={handleDownloadBillingRecordsExcel}
                disabled={filteredSubmittedBillingRecords.length === 0}
              >
                Download Excel
              </button>
            </div>

            {filteredSubmittedBillingRecords.length === 0 ? (
          <div className="billings-page__verification-empty">
            No billing records have been submitted yet.
          </div>
        ) : (
          <>
            <div
              className="billings-page__billing-records-table-wrapper"
              style={{
                width: "100%",
                overflowX: "hidden",
              }}
            >
              <table
                className="billings-page__billing-records-table"
                style={{
                  width: "100%",
                  minWidth: 0,
                  maxWidth: "100%",
                  tableLayout: "fixed",
                }}
              >
                <colgroup>
                  <col style={{ width: "110px" }} />
                  <col style={{ width: "105px" }} />
                  <col style={{ width: "105px" }} />
                  <col />
                  <col />
                  <col style={{ width: "120px" }} />
                  <col style={{ width: "130px" }} />
                  <col style={{ width: "92px" }} />
                </colgroup>
                <thead>
                  <tr>
                    {[
                      ["Customer Code", "billingCode"],
                      ["Complex Code", "complexCode"],
                      ["Screen Code", "screenCode"],
                      ["Screen Name", "screenName"],
                      ["Location", "location"],
                      ["Screen Fee", "effectiveScreenFee"],
                      ["Status", "status"],
                    ].map(([label, key]) => (
                      <SortableHeader key={key} label={label} sortKey={key} activeKey={billingRecordsSortKey} direction={billingRecordsSortOrder} onSort={(key) => toggleColumnSort(billingRecordsSortKey, setBillingRecordsSortKey, billingRecordsSortOrder, setBillingRecordsSortOrder, key)} style={{ whiteSpace: "normal", lineHeight: 1.2, padding: "10px 12px", boxSizing: "border-box", textAlign: "left", verticalAlign: "middle" }} />
                    ))}
                    <th style={{ whiteSpace: "normal", lineHeight: 1.2, padding: "10px 12px", boxSizing: "border-box", textAlign: "left", verticalAlign: "middle" }}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {pagedSubmittedBillingRecords.map((record) => {
                    const isActive =
                      activeBillingRecordId === record.billingRecordId;

                    return (
                      <tr
                        key={record.billingRecordId}
                        className={
                          isActive
                            ? "billings-page__billing-records-table-row billings-page__billing-records-table-row--active"
                            : "billings-page__billing-records-table-row"
                        }
                        style={
                          record.reasonPendingOverdue
                            ? {
                                background: "rgba(245, 158, 11, 0.16)",
                                boxShadow:
                                  "inset 4px 0 0 rgba(245, 158, 11, 0.85)",
                              }
                            : undefined
                        }
                      >
                        <td
                          title={record.billingCode || "-"}
                          style={{
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                            padding: "10px 12px",
                            boxSizing: "border-box",
                            verticalAlign: "middle",
                          }}
                        >
                          {record.billingCode || "-"}
                        </td>
                        <td
                          title={getDisplayComplexCode(record.complexCode)}
                          style={{
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                            padding: "10px 12px",
                            boxSizing: "border-box",
                            verticalAlign: "middle",
                          }}
                        >
                          {getDisplayComplexCode(record.complexCode)}
                        </td>
                        <td
                          title={record.screenCode || record.siteScope || "-"}
                          style={{
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                            padding: "10px 12px",
                            boxSizing: "border-box",
                            verticalAlign: "middle",
                          }}
                        >
                          {record.screenCode || record.siteScope || "-"}
                        </td>
                        <td
                          title={record.screenName || "-"}
                          style={{
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                            padding: "10px 12px",
                            boxSizing: "border-box",
                            verticalAlign: "middle",
                          }}
                        >
                          {record.screenName || "-"}
                        </td>
                        <td
                          title={record.location || "-"}
                          style={{
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                            padding: "10px 12px",
                            boxSizing: "border-box",
                            verticalAlign: "middle",
                          }}
                        >
                          {record.location || "-"}
                        </td>
                        <td
                          title={
                            getBillingRecordEffectiveScreenFee(
                              record,
                              commonComplexAllocations,
                            ) || "-"
                          }
                          style={{
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                            padding: "10px 12px",
                            boxSizing: "border-box",
                            verticalAlign: "middle",
                          }}
                        >
                          {formatInrAmount(
                            getBillingRecordEffectiveScreenFee(
                              record,
                              commonComplexAllocations,
                            ),
                          )}
                        </td>
                        <td
                          title={
                            record.reasonPendingOverdue
                              ? "Inactive — Reason Pending for more than 6 hours"
                              : record.displayBillingStatus || "Active"
                          }
                          style={{
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                            padding: "10px 12px",
                            boxSizing: "border-box",
                            verticalAlign: "middle",
                            fontWeight: record.reasonPendingOverdue ? 700 : 600,
                          }}
                        >
                          {record.reasonPendingOverdue
                            ? "Inactive — Reason Pending"
                            : record.displayBillingStatus || "Active"}
                        </td>
                        <td
                          className="billings-page__billing-records-table-action-cell"
                          style={{
                            width: "100%",
                            minWidth: 0,
                            maxWidth: "100%",
                            overflow: "visible",
                            boxSizing: "border-box",
                            padding: "8px 10px",
                            verticalAlign: "middle",
                          }}
                        >
                          <select
                            value=""
                            onChange={(event) => {
                              const action = event.target.value;

                              if (!action) {
                                return;
                              }

                              if (action === "view-details") {
                                handleOpenBillingRecord(record.billingRecordId);
                              } else if (action === "edit") {
                                handleEditBillingRecord(record.billingRecordId);
                              } else if (action === "price-change") {
                                handleInitiatePriceChangeFromBillingRecord(record);
                              }

                              event.target.value = "";
                            }}
                            aria-label={`Action for ${
                              record.screenCode || record.siteScope || "billing record"
                            }`}
                            title="Select Action"
                            style={{
                              width: "100%",
                              minWidth: 0,
                              maxWidth: "100%",
                              height: 34,
                              boxSizing: "border-box",
                              background: "transparent",
                              color: "inherit",
                              border: "1px solid rgba(96, 165, 250, 0.35)",
                              borderRadius: 8,
                              padding: "0 8px",
                            }}
                          >
                            <option
                              value=""
                              style={{ color: "#0f172a", backgroundColor: "#ffffff" }}
                            >
                              Select
                            </option>
                            <option
                              value="view-details"
                              style={{ color: "#0f172a", backgroundColor: "#ffffff" }}
                            >
                              View Details
                            </option>
                            <option
                              value="edit"
                              style={{ color: "#0f172a", backgroundColor: "#ffffff" }}
                            >
                              Edit
                            </option>
                            <option
                              value="price-change"
                              style={{ color: "#0f172a", backgroundColor: "#ffffff" }}
                            >
                              Price Change &amp; Reallocation
                            </option>
                          </select>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, marginTop: 12, flexWrap: "wrap" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                <span className="billings-page__helper" style={{ margin: 0 }}>
                  Showing {pagedSubmittedBillingRecords.length} of {filteredSubmittedBillingRecords.length} records
                </span>
                <div className="billings-page__field" style={{ marginBottom: 0, minWidth: 110 }}>
                                    <select
                    id="stage3-billing-records-rpp-bottom"
                    value={billingRecordsRowsPerPage}
                    onChange={(event) => setBillingRecordsRowsPerPage(Number(event.target.value))}
                  >
                    {rowsPerPageOptions.map((option) => (
                      <option key={option} value={option}>{option}</option>
                    ))}
                  </select>
                </div>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <button
                  type="button"
                  className="billings-page__secondary-button"
                  onClick={() => setBillingRecordsPage((page) => Math.max(1, page - 1))}
                  disabled={billingRecordsPage <= 1}
                >
                  Previous
                </button>
                <span>Page {billingRecordsPage} of {billingRecordsPageCount}</span>
                <button
                  type="button"
                  className="billings-page__secondary-button"
                  onClick={() => setBillingRecordsPage((page) => Math.min(billingRecordsPageCount, page + 1))}
                  disabled={billingRecordsPage >= billingRecordsPageCount}
                >
                  Next
                </button>
              </div>
            </div>

            {billingRecordDraft && activeBillingRecord && (
              <section className="billings-page__billing-record-details">
                <div className="billings-page__billing-record-details-header">
                  <div>
                    <h2>
                      {billingRecordEditMode
                        ? "Edit Billing Record"
                        : "Billing Record Details"}
                    </h2>
                    <span className="billings-page__billing-section-text">
                      {billingRecordEditMode
                        ? "Billing INC internal remarks can be edited here. Temporary monthly invoice and receipt/payment fallback is handled from ERP Monthly Billing & Payment."
                        : "ERP financial history is read-only by default. Until ERP sync starts, Invoice Number, Invoice Date, and receipt/payment values can be entered manually from the monthly Edit action."}
                    </span>
                  </div>

                  <button
                    type="button"
                    className="billings-page__secondary-button"
                    onClick={handleCloseBillingRecord}
                  >
                    Close
                  </button>
                </div>

                <section className="billings-page__billing-section">
                  <div className="billings-page__billing-section-heading">
                    <h2>Billing Reference</h2>
                  </div>

                  <div
                    className="billings-page__field-grid"
                    style={{
                      display: "grid",
                      gridTemplateColumns: "repeat(4, minmax(0, 1fr))",
                      gap: 12,
                      alignItems: "end",
                    }}
                  >
                    <div className="billings-page__field">
                      <label>Customer Code</label>
                      <input
                        type="text"
                        value={billingRecordDraft.billingCode || ""}
                        readOnly
                      />
                    </div>

                    <div className="billings-page__field">
                      <label>Screen Code</label>
                      <input
                        type="text"
                        value={
                          activeBillingRecord.screenCode ||
                          billingRecordDraft.siteScope ||
                          ""
                        }
                        readOnly
                      />
                    </div>

                    <div className="billings-page__field">
                      <label>Screen Name</label>
                      <input
                        type="text"
                        value={activeBillingRecord.screenName || ""}
                        readOnly
                      />
                    </div>

                    <div className="billings-page__field">
                      <label>Location</label>
                      <input
                        type="text"
                        value={activeBillingRecord.location || ""}
                        readOnly
                      />
                    </div>

                    <div className="billings-page__field">
                      <label>Subscription Mode</label>
                      <input
                        type="text"
                        value={activeBillingRecord.subscriptionMode || ""}
                        readOnly
                      />
                    </div>

                    <div className="billings-page__field">
                      <label>Subscription Fee</label>
                      <input
                        type="text"
                        value={getBillingRecordEffectiveScreenFee(
                          activeBillingRecord,
                          commonComplexAllocations,
                        )}
                        readOnly
                      />
                    </div>

                    <div className="billings-page__field">
                      <label>Billing Mode</label>
                      <input
                        type="text"
                        value={billingRecordDraft.billingMode || ""}
                        readOnly
                      />
                    </div>

                    <div className="billings-page__field">
                      <label>Billing Status</label>
                      {billingRecordEditMode ? (
                        <select
                          value={billingRecordDraft.billingStatus || "Active"}
                          onChange={(event) =>
                            handleBillingStatusDraftChange(event.target.value)
                          }
                          disabled={
                            (billingRecordDraft.billingStatus || "Active") !==
                              "Active" &&
                            getStage3RecordForBillingRecord(activeBillingRecord)
                              ?.billingLifecycleStatus === "Inactive"
                          }
                        >
                          <option value="Active">Active</option>
                          <option value="Inactive">Inactive</option>
                        </select>
                      ) : (
                        <input
                          type="text"
                          value={
                            getStage3RecordForBillingRecord(activeBillingRecord)
                              ?.billingLifecycleReason ||
                            getStage3RecordForBillingRecord(activeBillingRecord)
                              ?.billingLifecycleStatus ||
                            "Active"
                          }
                          readOnly
                        />
                      )}
                    </div>
                  </div>
                </section>

                <section
                  className="billings-page__billing-section"
                  style={
                    isErpMonthlyFullscreen
                      ? {
                          position: "fixed",
                          inset: 12,
                          zIndex: 1200,
                          margin: 0,
                          padding: 18,
                          overflow: "auto",
                          background: "#081a2b",
                          border: "1px solid rgba(96, 165, 250, 0.40)",
                          borderRadius: 14,
                          boxShadow: "0 20px 60px rgba(0, 0, 0, 0.45)",
                        }
                      : undefined
                  }
                >
                  {isErpMonthlyFullscreen ? (
                    <div
                      style={{
                        position: "sticky",
                        top: 0,
                        zIndex: 10,
                        display: "flex",
                        justifyContent: "flex-end",
                        marginBottom: 8,
                      }}
                    >
                      <button
                        type="button"
                        className="billings-page__secondary-button"
                        onClick={() => setIsErpMonthlyFullscreen(false)}
                        style={{
                          minWidth: 110,
                          boxShadow: "0 6px 18px rgba(0,0,0,0.30)",
                        }}
                      >
                        Minimize
                      </button>
                    </div>
                  ) : null}

                  <div
                    className="billings-page__billing-section-heading"
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      gap: 16,
                      alignItems: "end",
                      flexWrap: "wrap",
                    }}
                  >
                    <div>
                      <h2>ERP Monthly Billing &amp; Payment</h2>
                      <span className="billings-page__billing-section-text">
                        Monthly rows are shown from Billing Start Date through the
                        current month. ERP invoice, receipt, and payment values are
                        overlaid when available.
                      </span>
                    </div>

                    <div
                      style={{
                        display: "flex",
                        gap: 12,
                        alignItems: "end",
                        flexWrap: "wrap",
                      }}
                    >
                      <div
                        className="billings-page__field"
                        style={{ minWidth: 170, marginBottom: 0 }}
                      >
                        <label htmlFor="billing-record-financial-year">
                          Financial Year
                        </label>
                        <select
                          id="billing-record-financial-year"
                          value={billingRecordFinancialYear}
                          onChange={(event) => {
                            setBillingRecordFinancialYear(event.target.value);
                            setBillingRecordMonthFilter("All");
                            setBillingRecordSelectedEntryKey("");
                          }}
                        >
                          {billingRecordFinancialYearOptions.map((year) => (
                            <option key={year} value={year}>
                              {year}
                            </option>
                          ))}
                          <option value="All">All Years</option>
                        </select>
                      </div>

                      <div
                        className="billings-page__field"
                        style={{ minWidth: 170, marginBottom: 0 }}
                      >
                        <label htmlFor="billing-record-month-filter">
                          Billing Month
                        </label>
                        <select
                          id="billing-record-month-filter"
                          value={billingRecordMonthFilter}
                          onChange={(event) => {
                            setBillingRecordMonthFilter(event.target.value);
                            setBillingRecordSelectedEntryKey("");
                          }}
                        >
                          <option value="All">All Months</option>
                          {billingRecordMonthOptions.map((month) => (
                            <option key={month.value} value={month.value}>
                              {month.label}
                            </option>
                          ))}
                        </select>
                      </div>

                      <button
                        type="button"
                        className="billings-page__secondary-button"
                        onClick={() =>
                          setIsErpMonthlyFullscreen(
                            isErpMonthlyFullscreen ? false : true,
                          )
                        }
                        style={{ minWidth: 110 }}
                      >
                        {isErpMonthlyFullscreen ? "Minimize" : "Expand"}
                      </button>
                    </div>
                  </div>

                  {filteredBillingRecordFinancialRows.length === 0 ? (
                    <div className="billings-page__verification-empty">
                      No ERP billing/payment records are available for the selected
                      Financial Year.
                    </div>
                  ) : (
                    <div
                      className="billings-page__billing-records-table-wrapper"
                      style={{ width: "100%", overflowX: "hidden" }}
                    >
                      <table
                        className="billings-page__billing-records-table"
                        style={{
                          width: "100%",
                          minWidth: 0,
                          maxWidth: "100%",
                          tableLayout: "fixed",
                        }}
                      >
                        <colgroup>
                          <col style={{ width: "14%" }} />
                          <col style={{ width: "24%" }} />
                          <col style={{ width: "14%" }} />
                          <col style={{ width: "14%" }} />
                          <col style={{ width: "18%" }} />
                          <col style={{ width: "16%" }} />
                        </colgroup>
                        <thead>
                          <tr>
                            <th>Billing Month</th>
                            <th>Billing Period</th>
                            <th>Invoice No.</th>
                            <th>Receipt No.</th>
                            <th>Payment Status</th>
                            <th>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filteredBillingRecordFinancialRows.map((row) => (
                            <tr key={row.rowKey}>
                              <td style={{ whiteSpace: "normal", wordBreak: "break-word" }}>
                                {row.billingMonth || "-"}
                              </td>
                              <td style={{ whiteSpace: "normal", wordBreak: "break-word" }}>
                                {row.lifecycleBillable === false
                                  ? row.lifecyclePeriodLabel || "Paused / Not Billable"
                                  : formatBillingPeriod(row.periodFrom, row.periodTo)}
                              </td>
                              <td style={{ whiteSpace: "normal", wordBreak: "break-word" }}>
                                {row.invoiceNumber || "-"}
                              </td>
                              <td style={{ whiteSpace: "normal", wordBreak: "break-word" }}>
                                {row.receiptNumber || "-"}
                              </td>
                              <td style={{ whiteSpace: "normal", wordBreak: "break-word" }}>
                                {row.paymentStatus || "Pending"}
                              </td>
                              <td style={{ minWidth: 0 }}>
                                {normalizeValue(row.submittedPeriodTreatment).toLowerCase() ===
                                "waived" ? (
                                  <span
                                    title="Waived First Time billing period"
                                    aria-label="Waived First Time billing period"
                                  >
                                    -
                                  </span>
                                ) : (
                                  <select
                                    value=""
                                    onChange={(event) => {
                                      const action = event.target.value;

                                      if (!action) {
                                        return;
                                      }

                                      openBillingRecordMonthDetails(row, action);
                                      event.target.value = "";
                                    }}
                                    aria-label={`Action for ${row.billingMonth || "billing month"}`}
                                    style={{
                                      width: "100%",
                                      minWidth: 0,
                                      maxWidth: "100%",
                                      height: 34,
                                      boxSizing: "border-box",
                                      background: "transparent",
                                      color: "inherit",
                                      border: "1px solid rgba(96, 165, 250, 0.35)",
                                      borderRadius: 8,
                                      padding: "0 8px",
                                    }}
                                  >
                                    <option
                                      value=""
                                      style={{ color: "#0f172a", backgroundColor: "#ffffff" }}
                                    >
                                      Select
                                    </option>
                                    <option
                                      value="view"
                                      style={{ color: "#0f172a", backgroundColor: "#ffffff" }}
                                    >
                                      View
                                    </option>
                                    <option
                                      value="edit"
                                      style={{ color: "#0f172a", backgroundColor: "#ffffff" }}
                                    >
                                      Edit
                                    </option>
                                  </select>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}

                  {selectedBillingRecordFinancialRow && (
                    <section
                      className="billings-page__billing-section"
                      style={{ marginTop: 18 }}
                    >
                      <div className="billings-page__billing-section-heading">
                        <h2>
                          {selectedBillingRecordFinancialRow.billingMonth} - ERP
                          Billing Details
                        </h2>
                        <span className="billings-page__billing-section-text">
                          {billingRecordMonthDetailMode === "edit"
                            ? "Manual fallback entry is available only for missing receipt/payment details when ERP sync has not populated them."
                            : "Read-only monthly invoice, receipt, and payment details."}
                        </span>
                      </div>

                      <div className="billings-page__field-grid">
                        <div className="billings-page__field">
                          <label>Billing Period</label>
                          <input
                            type="text"
                            value={
                              isSelectedBillingMonthLifecycleLocked
                                ? selectedBillingRecordFinancialRow.lifecyclePeriodLabel || "Paused / Not Billable"
                                : formatBillingPeriod(
                                    selectedBillingRecordFinancialRow.periodFrom,
                                    selectedBillingRecordFinancialRow.periodTo,
                                  )
                            }
                            readOnly
                          />
                        </div>

                        <div className="billings-page__field">
                          <label>Billing Amount</label>
                          <input
                            type="text"
                            value={
                              selectedBillingRecordFinancialRow.billingAmount ||
                              ""
                            }
                            readOnly
                          />
                        </div>

                        <div className="billings-page__field">
                          <label>Invoice Number</label>
                          <input
                            type="text"
                            value={
                              billingRecordMonthDetailMode === "edit"
                                ? billingRecordManualPaymentDraft?.invoiceNumber || ""
                                : selectedBillingRecordFinancialRow.invoiceNumber || ""
                            }
                            onChange={(event) =>
                              updateBillingRecordManualPaymentDraft(
                                "invoiceNumber",
                                event.target.value,
                              )
                            }
                            readOnly={billingRecordMonthDetailMode !== "edit" || isSelectedBillingMonthLifecycleLocked}
                          />
                        </div>

                        <div className="billings-page__field">
                          <label>Invoice Date</label>
                          <input
                            type={billingRecordMonthDetailMode === "edit" && !isSelectedBillingMonthLifecycleLocked ? "date" : "text"}
                            value={
                              billingRecordMonthDetailMode === "edit"
                                ? billingRecordManualPaymentDraft?.invoiceDate || ""
                                : selectedBillingRecordFinancialRow.invoiceDate || ""
                            }
                            onChange={(event) =>
                              updateBillingRecordManualPaymentDraft(
                                "invoiceDate",
                                event.target.value,
                              )
                            }
                            readOnly={billingRecordMonthDetailMode !== "edit" || isSelectedBillingMonthLifecycleLocked}
                          />
                        </div>

                        <div className="billings-page__field">
                          <label>Invoice Amount Before GST</label>
                          <input
                            type="text"
                            inputMode="decimal"
                            value={
                              billingRecordMonthDetailMode === "edit"
                                ? billingRecordManualPaymentDraft
                                    ?.invoiceAmountBeforeGST || ""
                                : selectedBillingRecordFinancialRow
                                    .invoiceAmountBeforeGST || ""
                            }
                            onChange={(event) =>
                              updateBillingRecordManualPaymentDraft(
                                "invoiceAmountBeforeGST",
                                sanitizeAmount(event.target.value),
                              )
                            }
                            readOnly={billingRecordMonthDetailMode !== "edit" || isSelectedBillingMonthLifecycleLocked}
                          />
                        </div>

                        <div className="billings-page__field">
                          <label>Receipt Number</label>
                          <input
                            type="text"
                            value={
                              billingRecordMonthDetailMode === "edit"
                                ? billingRecordManualPaymentDraft?.receiptNumber || ""
                                : selectedBillingRecordFinancialRow.receiptNumber || ""
                            }
                            onChange={(event) =>
                              updateBillingRecordManualPaymentDraft(
                                "receiptNumber",
                                event.target.value,
                              )
                            }
                            readOnly={billingRecordMonthDetailMode !== "edit" || isSelectedBillingMonthLifecycleLocked}
                          />
                        </div>

                        <div className="billings-page__field">
                          <label>Receipt Date</label>
                          <input
                            type={billingRecordMonthDetailMode === "edit" && !isSelectedBillingMonthLifecycleLocked ? "date" : "text"}
                            value={
                              billingRecordMonthDetailMode === "edit"
                                ? billingRecordManualPaymentDraft?.receiptDate || ""
                                : selectedBillingRecordFinancialRow.receiptDate || ""
                            }
                            onChange={(event) =>
                              updateBillingRecordManualPaymentDraft(
                                "receiptDate",
                                event.target.value,
                              )
                            }
                            readOnly={billingRecordMonthDetailMode !== "edit" || isSelectedBillingMonthLifecycleLocked}
                          />
                        </div>

                        <div className="billings-page__field">
                          <label>Received Amount</label>
                          <input
                            type="text"
                            inputMode="decimal"
                            value={
                              billingRecordMonthDetailMode === "edit"
                                ? billingRecordManualPaymentDraft?.receivedAmount || ""
                                : selectedBillingRecordFinancialRow.receivedAmount || ""
                            }
                            onChange={(event) =>
                              updateBillingRecordManualPaymentDraft(
                                "receivedAmount",
                                event.target.value.replace(/[^0-9.]/g, ""),
                              )
                            }
                            readOnly={billingRecordMonthDetailMode !== "edit" || isSelectedBillingMonthLifecycleLocked}
                          />
                        </div>

                        <div className="billings-page__field">
                          <label>Payment Status</label>
                          {billingRecordMonthDetailMode === "edit" && !isSelectedBillingMonthLifecycleLocked ? (
                            <select
                              value={
                                billingRecordManualPaymentDraft?.paymentStatus ||
                                "Pending"
                              }
                              onChange={(event) =>
                                updateBillingRecordManualPaymentDraft(
                                  "paymentStatus",
                                  event.target.value,
                                )
                              }
                            >
                              <option value="Pending">Pending</option>
                              <option value="Partially Paid">Partially Paid</option>
                              <option value="Paid">Paid</option>
                              <option value="Disputed">Disputed</option>
                            </select>
                          ) : (
                            <input
                              type="text"
                              value={
                                selectedBillingRecordFinancialRow.paymentStatus ||
                                "Pending"
                              }
                              readOnly
                            />
                          )}
                        </div>

                        <div className="billings-page__field">
                          <label>Payment Remarks</label>
                          <textarea
                            rows={3}
                            value={
                              isSelectedBillingMonthLifecycleLocked
                                ? selectedBillingRecordFinancialRow.lifecycleRemarks ||
                                  selectedBillingRecordFinancialRow.paymentRemarks ||
                                  ""
                                : billingRecordMonthDetailMode === "edit"
                                  ? billingRecordManualPaymentDraft?.paymentRemarks || ""
                                  : selectedBillingRecordFinancialRow.paymentRemarks || ""
                            }
                            onChange={(event) =>
                              updateBillingRecordManualPaymentDraft(
                                "paymentRemarks",
                                event.target.value,
                              )
                            }
                            readOnly={billingRecordMonthDetailMode !== "edit" || isSelectedBillingMonthLifecycleLocked}
                          />
                        </div>
                      </div>

                      {selectedBillingRecordHasFuturePayment ? (
                        <div
                          className="billings-page__modal-notice"
                          style={{ marginTop: 12 }}
                          role="alert"
                        >
                          Future receipt date - payment not yet effective
                        </div>
                      ) : null}

                      {billingRecordMonthDetailMode === "edit" ? (
                        <div className="billings-page__modal-notice" style={{ marginTop: 12 }}>
                          {isSelectedBillingMonthLifecycleLocked
                            ? selectedBillingRecordFinancialRow.lifecycleRemarks ||
                              "This billing month falls completely within a paused/inactive period. Billing, invoice and payment entry are locked."
                            : "Temporary manual-entry fallback: Invoice Number, Invoice Date, and receipt/payment fields can be entered here until ERP synchronization is enabled. Billing Period and Billing Amount remain read-only. Saved manual entries are marked as Manual Entry in the billing record audit trail."}
                        </div>
                      ) : null}

                      {billingRecordManualPaymentMessage ? (
                        <div className="billings-page__modal-notice" style={{ marginTop: 12 }}>
                          {billingRecordManualPaymentMessage}
                        </div>
                      ) : null}

                      <div
                        className="billings-page__billing-record-details-actions"
                        style={{ marginTop: 12, display: "flex", gap: 10, justifyContent: "flex-end", flexWrap: "wrap" }}
                      >
                        {billingRecordMonthDetailMode === "edit" &&
                        !isSelectedBillingMonthLifecycleLocked ? (
                          <button
                            type="button"
                            className="billings-page__primary-button"
                            onClick={handleSaveManualBillingRecordPayment}
                          >
                            Save Manual Invoice / Payment
                          </button>
                        ) : null}

                        <button
                          type="button"
                          className="billings-page__secondary-button"
                          onClick={() => {
                            setBillingRecordSelectedEntryKey("");
                            setBillingRecordMonthDetailMode("view");
                            setBillingRecordManualPaymentDraft(null);
                            setBillingRecordManualPaymentMessage("");
                          }}
                        >
                          Close Month Details
                        </button>
                      </div>
                    </section>
                  )}
                </section>

                <section className="billings-page__billing-section">
                  <div className="billings-page__billing-section-heading">
                    <h2>Billing INC Internal Notes</h2>
                    <span className="billings-page__billing-section-text">
                      Internal notes are maintained in Billing INC. ERP financial
                      values above remain read-only.
                    </span>
                  </div>

                  <div className="billings-page__field">
                    <label>Billing Remarks</label>
                    <textarea
                      rows={3}
                      value={billingRecordDraft.billingRemarks || ""}
                      onChange={(event) =>
                        updateBillingRecordDraft(
                          "billingRemarks",
                          event.target.value,
                        )
                      }
                      readOnly={!billingRecordEditMode}
                      placeholder={
                        billingRecordEditMode
                          ? "Enter internal Billing INC remarks"
                          : "No internal remarks"
                      }
                    />
                  </div>

                  {billingRecordEditMode ? (
                    <div className="billings-page__billing-record-details-actions">
                      <button
                        type="button"
                        className="billings-page__primary-button"
                        onClick={handleSaveBillingInternalEdit}
                      >
                        Save Internal Edit
                      </button>
                    </div>
                  ) : null}

                  {billingRecordSaveMessage ? (
                    <div className="billings-page__modal-notice">
                      {billingRecordSaveMessage}
                    </div>
                  ) : null}
                </section>

                <section className="billings-page__billing-section">
                  <div className="billings-page__billing-section-heading">
                    <h2>Audit Information</h2>
                  </div>

                  <div className="billings-page__field-grid">
                    <div className="billings-page__field">
                      <label>Created At</label>
                      <input
                        type="text"
                        value={formatTimestampValue(billingRecordDraft.createdAt)}
                        readOnly
                      />
                    </div>

                    <div className="billings-page__field">
                      <label>Last Updated At</label>
                      <input
                        type="text"
                        value={formatTimestampValue(billingRecordDraft.lastUpdatedAt)}
                        readOnly
                      />
                    </div>
                  </div>
                </section>

              </section>
            )}
          </>
        )}
          </>
        )}
      </article>

      <section className="billings-page__closure">
        <div className="billings-page__closure-card">
          <div className="billings-page__closure-header">
            <div>
              <h2>Paused / Inactive</h2>
              <span className="billings-page__billing-section-text">
                Only sites moved out of Active billing are shown here. Update the Reason and effective date without changing historical Billing Records.
              </span>
            </div>

            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                marginLeft: "auto",
                flexShrink: 0,
              }}
            >
              <span className="billings-page__card-badge">
                {closureRows.length} Records
              </span>
              <button
                type="button"
                className="billings-page__secondary-button"
                onClick={() =>
                  handleBillingSectionToggle(billingSections.CLOSURE)
                }
              >
                {activeBillingSection === billingSections.CLOSURE
                  ? "Minimize"
                  : "Expand"}
              </button>
            </div>
          </div>

          {activeBillingSection === billingSections.CLOSURE && (
            <>
              <div
                className="billings-page__verification-actions"
                style={{
                  alignItems: "end",
                  flexWrap: "wrap",
                  marginBottom: 16,
                }}
              >
                <div
                  className="billings-page__field"
                  style={{ flex: "1 1 420px", marginBottom: 0 }}
                >
                  <label htmlFor="stage3-closure-search">Search</label>
                  <input
                    id="stage3-closure-search"
                    type="search"
                    value={closureSearchTerm}
                    onChange={(event) =>
                      setClosureSearchTerm(event.target.value)
                    }
                    placeholder="Search Customer Code, Complex Code, Screen Code, Screen Name, Location, or Status"
                  />
                </div>

              </div>

              {filteredClosureRows.length === 0 ? (
                <div className="billings-page__closure-empty">
                  No Paused / Inactive sites are available.
                </div>
              ) : (
                <div
                  className="billings-page__closure-table-wrapper"
                  style={{
                    width: "100%",
                    maxWidth: "100%",
                    overflowX: "hidden",
                  }}
                >
                  <table
                    className="billings-page__closure-table"
                    style={{
                      width: "100%",
                      minWidth: 0,
                      maxWidth: "100%",
                      tableLayout: "fixed",
                    }}
                  >
                    <colgroup>
                      <col style={{ width: "14%" }} />
                      <col style={{ width: "13%" }} />
                      <col style={{ width: "13%" }} />
                      <col style={{ width: "22%" }} />
                      <col style={{ width: "16%" }} />
                      <col style={{ width: "13%" }} />
                      <col style={{ width: "9%" }} />
                    </colgroup>
                    <thead>
                      <tr>
                        <SortableHeader label="Customer Code" sortKey="billingCode" activeKey={closureSortKey} direction={closureSortOrder} onSort={(key) => toggleColumnSort(closureSortKey, setClosureSortKey, closureSortOrder, setClosureSortOrder, key)} />
                        <SortableHeader label="Complex Code" sortKey="complexCode" activeKey={closureSortKey} direction={closureSortOrder} onSort={(key) => toggleColumnSort(closureSortKey, setClosureSortKey, closureSortOrder, setClosureSortOrder, key)} />
                        <SortableHeader label="Screen Code" sortKey="screenCode" activeKey={closureSortKey} direction={closureSortOrder} onSort={(key) => toggleColumnSort(closureSortKey, setClosureSortKey, closureSortOrder, setClosureSortOrder, key)} />
                        <SortableHeader label="Screen Name" sortKey="screenName" activeKey={closureSortKey} direction={closureSortOrder} onSort={(key) => toggleColumnSort(closureSortKey, setClosureSortKey, closureSortOrder, setClosureSortOrder, key)} />
                        <SortableHeader label="Location" sortKey="location" activeKey={closureSortKey} direction={closureSortOrder} onSort={(key) => toggleColumnSort(closureSortKey, setClosureSortKey, closureSortOrder, setClosureSortOrder, key)} />
                        <SortableHeader label="Current Billing Status" sortKey="currentBillingStatus" activeKey={closureSortKey} direction={closureSortOrder} onSort={(key) => toggleColumnSort(closureSortKey, setClosureSortKey, closureSortOrder, setClosureSortOrder, key)} />
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {pagedClosureRows.map((record) => (
                        <tr key={record.recordId}>
                          <td style={{ whiteSpace: "normal", overflowWrap: "anywhere" }}>
                            {record.billingCode || "-"}
                          </td>
                          <td
                            title={getDisplayComplexCode(record.complexCode)}
                            style={{ whiteSpace: "normal", overflowWrap: "anywhere" }}
                          >
                            {getDisplayComplexCode(record.complexCode)}
                          </td>
                          <td style={{ whiteSpace: "normal", overflowWrap: "anywhere" }}>
                            {record.screenCode || "-"}
                          </td>
                          <td style={{ whiteSpace: "normal", overflowWrap: "anywhere" }}>
                            {record.screenName || "-"}
                          </td>
                          <td style={{ whiteSpace: "normal", overflowWrap: "anywhere" }}>
                            {record.location || "-"}
                          </td>
                          <td style={{ whiteSpace: "normal", overflowWrap: "anywhere" }}>
                            {record.currentBillingStatus}
                          </td>
                          <td style={{ minWidth: 0 }}>
                            {["Billing Paused", "Site Inactive"].includes(
                              record.currentBillingStatus,
                            ) ? (
                              <button
                                type="button"
                                className="billings-page__secondary-button"
                                style={{
                                  width: "100%",
                                  minWidth: 0,
                                  paddingInline: 6,
                                }}
                                onClick={() =>
                                  handleOpenReactivate(record.recordId)
                                }
                              >
                                Reactivate
                              </button>
                            ) : (
                              <button
                                type="button"
                                className="billings-page__secondary-button"
                                style={{
                                  width: "100%",
                                  minWidth: 0,
                                  paddingInline: 6,
                                }}
                                onClick={() =>
                                  handleOpenClosureRecord(record.recordId)
                                }
                              >
                                Open
                              </button>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {filteredClosureRows.length > 0 && (
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    gap: 12,
                    marginTop: 12,
                    flexWrap: "wrap",
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 10,
                      flexWrap: "wrap",
                    }}
                  >
                    <span
                      className="billings-page__helper"
                      style={{ margin: 0 }}
                    >
                      Showing {pagedClosureRows.length} of{" "}
                      {filteredClosureRows.length} records
                    </span>

                    <div
                      className="billings-page__field"
                      style={{ marginBottom: 0, minWidth: 110 }}
                    >
                      <select
                        id="stage3-closure-rpp-bottom"
                        value={closureRowsPerPage}
                        onChange={(event) =>
                          setClosureRowsPerPage(Number(event.target.value))
                        }
                      >
                        {rowsPerPageOptions.map((option) => (
                          <option key={option} value={option}>
                            {option}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>

                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 8,
                    }}
                  >
                    <button
                      type="button"
                      className="billings-page__secondary-button"
                      onClick={() =>
                        setClosurePage((page) => Math.max(1, page - 1))
                      }
                      disabled={closurePage <= 1}
                    >
                      Previous
                    </button>
                    <span>
                      Page {closurePage} of {closurePageCount}
                    </span>
                    <button
                      type="button"
                      className="billings-page__secondary-button"
                      onClick={() =>
                        setClosurePage((page) =>
                          Math.min(closurePageCount, page + 1),
                        )
                      }
                      disabled={closurePage >= closurePageCount}
                    >
                      Next
                    </button>
                  </div>
                </div>
              )}

              {activeClosureRecord ? (
                <section
                  id="stage3-billing-closure-editor"
                  className="billings-page__billing-record-details"
                  style={{ marginTop: 20, width: "100%", boxSizing: "border-box" }}
                >
                  {closureEditorMode === "reactivate" ? (
                    <>
                      <div className="billings-page__billing-section-heading">
                        <div>
                          <h2>Reactivate Billing</h2>
                          <span className="billings-page__billing-section-text">
                            {activeClosureRecord.screenName || "-"} ·{" "}
                            {activeClosureRecord.screenCode || "-"}
                          </span>
                        </div>

                        <span className="billings-page__card-badge">
                          {activeClosureRecord.currentBillingStatus}
                        </span>
                      </div>

                      <div
                        className="billings-page__billing-record-reference-grid"
                        style={{
                          marginTop: 16,
                          gridTemplateColumns:
                            "repeat(auto-fit, minmax(240px, 1fr))",
                          width: "100%",
                        }}
                      >
                        <div className="billings-page__field">
                          <label>
                            {activeClosureRecord.currentBillingStatus ===
                            "Billing Paused"
                              ? "Paused From"
                              : "Inactive From"}
                          </label>
                          <input
                            type="date"
                            value={
                              getBillingPauseDate(activeClosureRecord)
                                ? formatLocalDateValue(
                                    getBillingPauseDate(activeClosureRecord),
                                  )
                                : ""
                            }
                            readOnly
                            aria-readonly="true"
                          />
                        </div>

                        <div className="billings-page__field">
                          <label htmlFor="billing-restore-active-from">
                            Active From
                          </label>
                          <input
                            id="billing-restore-active-from"
                            type="date"
                            min={
                              getBillingPauseDate(activeClosureRecord)
                                ? formatLocalDateValue(
                                    getBillingPauseDate(activeClosureRecord),
                                  )
                                : undefined
                            }
                            value={restoreActiveDraft.activeFrom}
                            onChange={(event) => {
                              setRestoreActiveDraft((currentDraft) => ({
                                ...currentDraft,
                                activeFrom: event.target.value,
                              }));
                              setClosureMessage("");
                            }}
                          />
                        </div>

                        <div className="billings-page__field">
                          <label htmlFor="billing-restore-active-remarks">
                            Restore Remarks
                          </label>
                          <textarea
                            id="billing-restore-active-remarks"
                            rows={3}
                            value={restoreActiveDraft.remarks}
                            onChange={(event) => {
                              setRestoreActiveDraft((currentDraft) => ({
                                ...currentDraft,
                                remarks: event.target.value,
                              }));
                              setClosureMessage("");
                            }}
                            placeholder="Enter remarks for restoring this site to Active"
                          />
                        </div>
                      </div>

                      <div
                        className="billings-page__billing-record-details"
                        style={{
                          marginTop: 20,
                          width: "100%",
                          boxSizing: "border-box",
                        }}
                      >
                        <div
                          className="billings-page__billing-section-heading"
                          style={{ marginBottom: 12 }}
                        >
                          <div>
                            <h3 style={{ margin: 0 }}>Status History</h3>
                            <span className="billings-page__billing-section-text">
                              Most recent status change is shown first.
                            </span>
                          </div>
                        </div>

                        {getSortedClosureHistory(activeClosureRecord).length ? (
                          <div
                            className="billings-page__closure-table-wrapper"
                            style={{
                              width: "100%",
                              maxWidth: "100%",
                              overflowX: "auto",
                            }}
                          >
                            <table
                              className="billings-page__closure-table"
                              style={{
                                width: "100%",
                                minWidth: 760,
                                tableLayout: "fixed",
                              }}
                            >
                              <colgroup>
                                <col style={{ width: "17%" }} />
                                <col style={{ width: "17%" }} />
                                <col style={{ width: "17%" }} />
                                <col style={{ width: "25%" }} />
                                <col style={{ width: "12%" }} />
                                <col style={{ width: "12%" }} />
                              </colgroup>
                              <thead>
                                <tr>
                                  <th>Status</th>
                                  <th>Effective From</th>
                                  <th>Active From</th>
                                  <th>Remarks / Reason</th>
                                  <th>Updated By</th>
                                  <th>Updated At</th>
                                </tr>
                              </thead>
                              <tbody>
                                {getSortedClosureHistory(activeClosureRecord).map(
                                  (historyItem, historyIndex) => {
                                    const historyStatus = normalizeValue(
                                      historyItem?.newStatus,
                                    );
                                    const isActiveHistory =
                                      historyStatus.toLowerCase() === "active";

                                    const effectiveFrom = normalizeValue(
                                      historyItem?.effectiveDate,
                                    );

                                    const activeFrom = isActiveHistory
                                      ? effectiveFrom
                                      : normalizeValue(
                                          historyItem?.activeFrom ||
                                            historyItem?.restoredFrom ||
                                            historyItem?.billingActiveFromDate,
                                        );

                                    return (
                                      <tr
                                        key={
                                          historyItem?.eventId ||
                                          `${historyStatus}-${effectiveFrom}-${historyIndex}`
                                        }
                                      >
                                        <td>
                                          {historyStatus ||
                                            normalizeValue(
                                              historyItem?.status,
                                            ) ||
                                            "-"}
                                        </td>
                                        <td>
                                          {isActiveHistory
                                            ? "-"
                                            : effectiveFrom || "-"}
                                        </td>
                                        <td>
                                          {activeFrom || "-"}
                                        </td>
                                        <td
                                          style={{
                                            whiteSpace: "normal",
                                            overflowWrap: "anywhere",
                                          }}
                                        >
                                          {normalizeValue(
                                            historyItem?.remarks ||
                                              historyItem?.reason ||
                                              historyItem?.restoreRemarks,
                                          ) || "-"}
                                        </td>
                                        <td>
                                          {normalizeValue(
                                            historyItem?.updatedBy ||
                                              historyItem?.changedBy,
                                          ) || "-"}
                                        </td>
                                        <td>
                                          {normalizeValue(
                                            historyItem?.updatedAt ||
                                              historyItem?.changedAt,
                                          ) || "-"}
                                        </td>
                                      </tr>
                                    );
                                  },
                                )}
                              </tbody>
                            </table>
                          </div>
                        ) : (
                          <div
                            className="billings-page__modal-notice"
                            style={{ marginTop: 8 }}
                          >
                            No previous Paused / Inactive status history is
                            available for this site yet.
                          </div>
                        )}
                      </div>

                      {closureMessage ? (
                        <div
                          className="billings-page__modal-notice"
                          style={{ marginTop: 12 }}
                        >
                          {closureMessage}
                        </div>
                      ) : null}

                      <div
                        className="billings-page__billing-record-details-actions"
                        style={{
                          marginTop: 14,
                          display: "flex",
                          gap: 10,
                          justifyContent: "flex-end",
                          flexWrap: "wrap",
                        }}
                      >
                        <button
                          type="button"
                          className="billings-page__secondary-button"
                          onClick={() => {
                            setClosureEditorMode("reason");
                            setActiveClosureRecordId("");
                            setRestoreActiveDraft({
                              activeFrom: "",
                              remarks: "",
                            });
                            setClosureMessage("");
                          }}
                        >
                          Close
                        </button>

                        <button
                          type="button"
                          className="billings-page__primary-button"
                          onClick={() =>
                            handleRestoreToActive(activeClosureRecord.recordId)
                          }
                        >
                          Save
                        </button>
                      </div>
                    </>
                  ) : (
                    <>
                      <div className="billings-page__billing-section-heading">
                        <div>
                          <h2>Site Details</h2>
                          <span className="billings-page__billing-section-text">
                            {activeClosureRecord.screenName || "-"} ·{" "}
                            {activeClosureRecord.screenCode || "-"}
                          </span>
                        </div>

                        <span className="billings-page__card-badge">
                          {activeClosureRecord.currentBillingStatus}
                        </span>
                      </div>

                      <div
                        className="billings-page__billing-record-reference-grid"
                        style={{
                          gridTemplateColumns:
                            "repeat(auto-fit, minmax(210px, 1fr))",
                          width: "100%",
                        }}
                      >
                        <div className="billings-page__field">
                          <label>Customer Code</label>
                          <input
                            type="text"
                            value={activeClosureRecord.billingCode || ""}
                            readOnly
                          />
                        </div>

                        <div className="billings-page__field">
                          <label>Complex Code</label>
                          <input
                            type="text"
                            value={getDisplayComplexCode(
                              activeClosureRecord.complexCode,
                            )}
                            readOnly
                          />
                        </div>

                        <div className="billings-page__field">
                          <label>Screen Code</label>
                          <input
                            type="text"
                            value={activeClosureRecord.screenCode || ""}
                            readOnly
                          />
                        </div>

                        <div className="billings-page__field">
                          <label>Location</label>
                          <input
                            type="text"
                            value={activeClosureRecord.location || ""}
                            readOnly
                          />
                        </div>
                      </div>

                      {["Site Removed", "Contract Closed"].includes(
                        activeClosureRecord.currentBillingStatus,
                      ) ? (
                        <div
                          className="billings-page__modal-notice"
                          style={{ marginTop: 12 }}
                        >
                          This status cannot be reactivated from Stage 3. If the
                          site rejoins, the workflow must start from the beginning.
                        </div>
                      ) : null}

                      <div
                        className="billings-page__billing-record-reference-grid"
                        style={{
                          marginTop: 16,
                          gridTemplateColumns:
                            "repeat(auto-fit, minmax(240px, 1fr))",
                          width: "100%",
                        }}
                      >
                        <div className="billings-page__field">
                          <label htmlFor="billing-closure-type">Reason</label>
                          <select
                            id="billing-closure-type"
                            value={closureDraft.closureType}
                            onChange={(event) =>
                              updateClosureDraft(
                                "closureType",
                                event.target.value,
                              )
                            }
                          >
                            <option value="">Select</option>
                            <option value="Billing Paused">Billing Paused</option>
                            <option value="Site Inactive">Site Inactive</option>
                            <option value="Site Removed">Site Removed</option>
                            <option value="Contract Closed">
                              Contract Quit / Closed
                            </option>
                          </select>
                        </div>

                        <div className="billings-page__field">
                          <label htmlFor="billing-closure-effective-date">
                            {closureDraft.closureType === "Billing Paused"
                              ? "Pause From"
                              : "Effective Date"}
                          </label>
                          <input
                            id="billing-closure-effective-date"
                            type="date"
                            value={closureDraft.effectiveDate}
                            onChange={(event) =>
                              updateClosureDraft(
                                "effectiveDate",
                                event.target.value,
                              )
                            }
                          />
                        </div>

                        <div className="billings-page__field">
                          <label htmlFor="billing-closure-remarks">
                            Remarks / Reason
                          </label>
                          <textarea
                            id="billing-closure-remarks"
                            rows={3}
                            value={closureDraft.remarks}
                            onChange={(event) =>
                              updateClosureDraft("remarks", event.target.value)
                            }
                            placeholder="Enter remarks for this status change"
                          />
                        </div>
                      </div>

                      {closureMessage ? (
                        <div
                          className="billings-page__modal-notice"
                          style={{ marginTop: 12 }}
                        >
                          {closureMessage}
                        </div>
                      ) : null}

                      <div
                        className="billings-page__billing-record-details-actions"
                        style={{
                          marginTop: 14,
                          display: "flex",
                          gap: 10,
                          justifyContent: "flex-end",
                          flexWrap: "wrap",
                        }}
                      >
                        <button
                          type="button"
                          className="billings-page__secondary-button"
                          onClick={() => {
                            setClosureEditorMode("reason");
                            setActiveClosureRecordId("");
                            setClosureDraft({
                              closureType: "",
                              effectiveDate: "",
                              remarks: "",
                            });
                            setClosureMessage("");
                          }}
                        >
                          Close
                        </button>

                        <button
                          type="button"
                          className="billings-page__primary-button"
                          onClick={handleSaveBillingClosure}
                        >
                          Save
                        </button>
                      </div>
                    </>
                  )}
                </section>
              ) : null}
            </>
          )}
        </div>
      </section>

        </>
      ) : activeWorkspace === workspaceModes.FOC ? (
        <article className="billings-page__card">
          <div className="billings-page__card-header">
            <div>
              <p className="billings-page__card-title">FoC</p>
              <span className="billings-page__card-badge billings-page__card-badge--editable">
                {focRows.length} Sites
              </span>
            </div>

          </div>

          <p className="billings-page__helper">
            Sites that entered Stage 3 as FoC remain in this section permanently.
            When converted to Billable, only the Current Commercial Status changes.
          </p>

          <div className="billings-page__verification-actions" style={{ alignItems: "end", flexWrap: "wrap", marginBottom: 16 }}>
            <div className="billings-page__field" style={{ flex: "1 1 360px", marginBottom: 0 }}>
              <label htmlFor="foc-search">Search FoC Sites</label>
              <input
                id="foc-search"
                type="search"
                value={focSearchTerm}
                onChange={(event) => setFocSearchTerm(event.target.value)}
                placeholder="Search Billing Name, Screen Code, Screen Name, City / Town or State"
              />
            </div>
            <div className="billings-page__field" style={{ minWidth: 145, marginBottom: 0 }}>
              <label htmlFor="foc-from-date">From Date</label>
              <input id="foc-from-date" type="date" value={focFromDate} onChange={(event) => setFocFromDate(event.target.value)} />
            </div>
            <div className="billings-page__field" style={{ minWidth: 145, marginBottom: 0 }}>
              <label htmlFor="foc-to-date">To Date</label>
              <input id="foc-to-date" type="date" value={focToDate} onChange={(event) => setFocToDate(event.target.value)} />
            </div>
          </div>

          {focRows.length === 0 ? (
            <div className="billings-page__verification-empty">
              No FoC sites with an eligible Billing Date are available yet.
            </div>
          ) : (
            <>
            <div className="billings-page__otf-table-wrapper billings-page__foc-table-wrapper">
              <table className="billings-page__otf-table billings-page__foc-table">
                <colgroup>
                  <col style={{ width: "16%" }} />
                  <col style={{ width: "20%" }} />
                  <col style={{ width: "16%" }} />
                  <col style={{ width: "14%" }} />
                  <col style={{ width: "12%" }} />
                  <col style={{ width: "22%" }} />
                </colgroup>
                <thead>
                  <tr>
                    <SortableHeader label="Screen Code" sortKey="screenCode" activeKey={focSortKey} direction={focSortOrder} onSort={(key) => toggleColumnSort(focSortKey, setFocSortKey, focSortOrder, setFocSortOrder, key)} />
                    <SortableHeader label="Screen Name" sortKey="erpScreenName" activeKey={focSortKey} direction={focSortOrder} onSort={(key) => toggleColumnSort(focSortKey, setFocSortKey, focSortOrder, setFocSortOrder, key)} />
                    <SortableHeader label="City / Town" sortKey="location" activeKey={focSortKey} direction={focSortOrder} onSort={(key) => toggleColumnSort(focSortKey, setFocSortKey, focSortOrder, setFocSortOrder, key)} />
                    <SortableHeader label="State" sortKey="state" activeKey={focSortKey} direction={focSortOrder} onSort={(key) => toggleColumnSort(focSortKey, setFocSortKey, focSortOrder, setFocSortOrder, key)} />
                    <SortableHeader label="Status" sortKey="commercialStatus" activeKey={focSortKey} direction={focSortOrder} onSort={(key) => toggleColumnSort(focSortKey, setFocSortKey, focSortOrder, setFocSortOrder, key)} />
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {pagedFocRows.map((record) => {
                    const currentStatus = getFoCCurrentCommercialStatus(record);
                    const isBillable = currentStatus === "Billable";
                    return (
                      <tr key={record.recordId}>
                        <td>{record.screenCode || "-"}</td>
                        <td>{record.erpScreenName || record.screenName || "-"}</td>
                        <td>{record.location || "-"}</td>
                        <td>{record.state || "-"}</td>
                        <td>
                          <span className="billings-page__card-badge">{currentStatus}</span>
                        </td>
                        <td className="billings-page__foc-action-cell">
                          <button
                            type="button"
                            className="billings-page__otf-open-button billings-page__foc-action-button"
                            onClick={() => handleOpenFocBillableEditor(record)}
                          >
                            {isBillable ? "Edit Billable" : "Change to Billable"}
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, marginTop: 12, flexWrap: "wrap" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                <span className="billings-page__helper" style={{ margin: 0 }}>
                  Showing {pagedFocRows.length} of {focRows.length} records
                </span>
                <div className="billings-page__field" style={{ marginBottom: 0, minWidth: 110 }}>
                  <select
                    id="stage3-foc-rpp-bottom"
                    aria-label="Rows per page"
                    value={focRowsPerPage}
                    onChange={(event) => setFocRowsPerPage(Number(event.target.value))}
                  >
                    {rowsPerPageOptions.map((option) => (
                      <option key={option} value={option}>{option}</option>
                    ))}
                  </select>
                </div>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <button
                  type="button"
                  className="billings-page__secondary-button"
                  onClick={() => setFocPage((page) => Math.max(1, page - 1))}
                  disabled={focPage <= 1}
                >
                  Previous
                </button>
                <span>Page {focPage} of {focPageCount}</span>
                <button
                  type="button"
                  className="billings-page__secondary-button"
                  onClick={() => setFocPage((page) => Math.min(focPageCount, page + 1))}
                  disabled={focPage >= focPageCount}
                >
                  Next
                </button>
              </div>
            </div>
            </>
          )}
        </article>
      ) : activeWorkspace === workspaceModes.EXPENSES ? (
        <article className="billings-page__card">
          <div className="billings-page__card-header">
            <div>
              <p className="billings-page__card-title">Other Expenses</p>
              <span className="billings-page__card-badge billings-page__card-badge--editable">
                Reference
              </span>
            </div>

          </div>

          <p className="billings-page__helper">
            Installation expenses captured in Site Information &amp; Installation.
            These amounts are for tracking only and are not billing adjustments.
          </p>

          <div className="billings-page__verification-actions" style={{ alignItems: "end", flexWrap: "wrap", marginBottom: 16 }}>
            <div className="billings-page__field" style={{ flex: "1 1 360px", marginBottom: 0 }}>
              <label htmlFor="other-expenses-search">Search Other Expenses</label>
              <input
                id="other-expenses-search"
                type="search"
                value={otherExpensesSearchTerm}
                onChange={(event) => setOtherExpensesSearchTerm(event.target.value)}
                placeholder="Search Customer Code, Complex Code, Screen Code / ID, Screen / Complex Name, Location, or State"
              />
            </div>
            <div className="billings-page__field" style={{ minWidth: 145, marginBottom: 0 }}>
              <label htmlFor="other-expenses-from-date">From Date</label>
              <input id="other-expenses-from-date" type="date" value={otherExpensesFromDate} onChange={(event) => setOtherExpensesFromDate(event.target.value)} />
            </div>
            <div className="billings-page__field" style={{ minWidth: 145, marginBottom: 0 }}>
              <label htmlFor="other-expenses-to-date">To Date</label>
              <input id="other-expenses-to-date" type="date" value={otherExpensesToDate} onChange={(event) => setOtherExpensesToDate(event.target.value)} />
            </div>
          </div>

          {filteredOtherExpenseRecords.length === 0 ? (
            <div className="billings-page__verification-empty">
              No Other Installation Expenses are available yet.
            </div>
          ) : (
            <>
              <div className="billings-page__otf-table-wrapper">
                <table className="billings-page__otf-table">
                  <thead>
                    <tr>
                      <SortableHeader label="Customer Code" sortKey="billingCode" activeKey={otherExpensesSortKey} direction={otherExpensesSortOrder} onSort={(key) => toggleColumnSort(otherExpensesSortKey, setOtherExpensesSortKey, otherExpensesSortOrder, setOtherExpensesSortOrder, key)} />
                      <SortableHeader label="Screen / Complex Name" sortKey="displayName" activeKey={otherExpensesSortKey} direction={otherExpensesSortOrder} onSort={(key) => toggleColumnSort(otherExpensesSortKey, setOtherExpensesSortKey, otherExpensesSortOrder, setOtherExpensesSortOrder, key)} />
                      <SortableHeader label="Location" sortKey="location" activeKey={otherExpensesSortKey} direction={otherExpensesSortOrder} onSort={(key) => toggleColumnSort(otherExpensesSortKey, setOtherExpensesSortKey, otherExpensesSortOrder, setOtherExpensesSortOrder, key)} />
                      <SortableHeader label="State" sortKey="state" activeKey={otherExpensesSortKey} direction={otherExpensesSortOrder} onSort={(key) => toggleColumnSort(otherExpensesSortKey, setOtherExpensesSortKey, otherExpensesSortOrder, setOtherExpensesSortOrder, key)} />
                      <SortableHeader label="Other Expenses" sortKey="otherExpenses" activeKey={otherExpensesSortKey} direction={otherExpensesSortOrder} onSort={(key) => toggleColumnSort(otherExpensesSortKey, setOtherExpensesSortKey, otherExpensesSortOrder, setOtherExpensesSortOrder, key)} />
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {pagedOtherExpenseRecords.map((record) => (
                      <tr
                        key={record.recordId}
                        className={
                          activeOtherExpenseRecord?.recordId === record.recordId
                            ? "billings-page__otf-table-row billings-page__otf-table-row--active"
                            : "billings-page__otf-table-row"
                        }
                      >
                        <td>{record.billingCode || "-"}</td>
                        <td>{record.displayName || "-"}</td>
                        <td>{record.location || "-"}</td>
                        <td>{record.state || "-"}</td>
                        <td>
                          ₹{Number(record.total || 0).toLocaleString("en-IN", {
                            minimumFractionDigits: 2,
                            maximumFractionDigits: 2,
                          })}
                        </td>
                        <td>
                          <button
                            type="button"
                            className="billings-page__otf-open-button"
                            onClick={() => handleOpenOtherExpenses(record.recordId)}
                          >
                            View
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, marginTop: 12, flexWrap: "wrap" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                  <span className="billings-page__helper" style={{ margin: 0 }}>
                    Showing {pagedOtherExpenseRecords.length} of {filteredOtherExpenseRecords.length} records
                  </span>
                  <div className="billings-page__field" style={{ marginBottom: 0, minWidth: 110 }}>
                    <select
                      id="stage3-other-expenses-rpp-bottom"
                      aria-label="Rows per page"
                      value={otherExpensesRowsPerPage}
                      onChange={(event) => setOtherExpensesRowsPerPage(Number(event.target.value))}
                    >
                      {rowsPerPageOptions.map((option) => (
                        <option key={option} value={option}>{option}</option>
                      ))}
                    </select>
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <button
                    type="button"
                    className="billings-page__secondary-button"
                    onClick={() => setOtherExpensesPage((page) => Math.max(1, page - 1))}
                    disabled={otherExpensesPage <= 1}
                  >
                    Previous
                  </button>
                  <span>Page {otherExpensesPage} of {otherExpensesPageCount}</span>
                  <button
                    type="button"
                    className="billings-page__secondary-button"
                    onClick={() => setOtherExpensesPage((page) => Math.min(otherExpensesPageCount, page + 1))}
                    disabled={otherExpensesPage >= otherExpensesPageCount}
                  >
                    Next
                  </button>
                </div>
              </div>

              {activeOtherExpenseRecord ? (
                <section className="billings-page__otf-details">
                  <div className="billings-page__billing-section-heading">
                    <h2>Other Expense Details</h2>
                    <span className="billings-page__billing-section-text">
                      {activeOtherExpenseRecord.displayName || "-"}
                    </span>
                  </div>

                  <div className="billings-page__otf-table-wrapper">
                    <table className="billings-page__otf-table">
                      <thead>
                        <tr>
                          <th>Screen</th>
                          <th>Expense Type</th>
                          <th>Amount Spent</th>
                          <th>Date of Capture</th>
                          <th>Remarks</th>
                        </tr>
                      </thead>
                      <tbody>
                        {activeOtherExpenseRecord.expenses.map((expense) => (
                          <tr key={expense.expenseId}>
                            <td>
                              {expense.screenCode || "-"}
                              {expense.screenName ? ` / ${expense.screenName}` : ""}
                            </td>
                            <td>{expense.expenseType || "-"}</td>
                            <td>
                              ₹{Number(expense.amount || 0).toLocaleString("en-IN", {
                                minimumFractionDigits: 2,
                                maximumFractionDigits: 2,
                              })}
                            </td>
                            <td>{formatBillingDateDisplay(expense.expenseDate)}</td>
                            <td>{expense.comments || "-"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <div
                    className="billings-page__helper"
                    style={{ marginTop: 12, fontWeight: 700 }}
                  >
                    Total Other Expenses: ₹
                    {Number(activeOtherExpenseRecord.total || 0).toLocaleString("en-IN", {
                      minimumFractionDigits: 2,
                      maximumFractionDigits: 2,
                    })}
                  </div>
                </section>
              ) : null}
            </>
          )}
        </article>
      ) : (
        <article className="billings-page__card billings-page__otf-card">
          <div className="billings-page__card-header">
            <div>
              <p className="billings-page__card-title">OTF Tracking</p>
              <span className="billings-page__card-badge billings-page__card-badge--editable">
                Active
              </span>
            </div>

          </div>

          <p className="billings-page__helper">
            Track applicable One Time Fee transactions and their invoice / RICBR references.
          </p>

          <div className="billings-page__verification-actions" style={{ alignItems: "end", flexWrap: "wrap", marginBottom: 16 }}>
            <div className="billings-page__field" style={{ flex: "1 1 360px", marginBottom: 0 }}>
              <label htmlFor="otf-search">Search OTF Transactions</label>
              <input
                id="otf-search"
                type="search"
                value={otfSearchTerm}
                onChange={(event) => setOtfSearchTerm(event.target.value)}
                placeholder="Search Billing ID, Complex Code, Screen Code / ID, Screen / Complex Name, City / Town, State, invoice, or RICBR"
              />
            </div>
            <div className="billings-page__field" style={{ minWidth: 145, marginBottom: 0 }}>
              <label htmlFor="otf-from-date">From Date</label>
              <input id="otf-from-date" type="date" value={otfFromDate} onChange={(event) => setOtfFromDate(event.target.value)} />
            </div>
            <div className="billings-page__field" style={{ minWidth: 145, marginBottom: 0 }}>
              <label htmlFor="otf-to-date">To Date</label>
              <input id="otf-to-date" type="date" value={otfToDate} onChange={(event) => setOtfToDate(event.target.value)} />
            </div>
          </div>

          {filteredOtfTransactions.length === 0 ? (
            <div className="billings-page__verification-empty">
              No OTF transactions are available yet.
            </div>
          ) : (
            <>
              <div className="billings-page__otf-table-wrapper">
                <table className="billings-page__otf-table">
                  <colgroup>
                    <col className="billings-page__otf-col-billing-code" />
                    <col className="billings-page__otf-col-screen-code" />
                    <col className="billings-page__otf-col-screen-name" />
                    <col className="billings-page__otf-col-city" />
                    <col className="billings-page__otf-col-state" />
                    <col className="billings-page__otf-col-action" />
                  </colgroup>
                  <thead>
                    <tr>
                      <SortableHeader label="Complex Code" sortKey="complexCode" activeKey={otfSortKey} direction={otfSortOrder} onSort={(key) => toggleColumnSort(otfSortKey, setOtfSortKey, otfSortOrder, setOtfSortOrder, key)} />
                      <SortableHeader label="Screen Code" sortKey="screenCode" activeKey={otfSortKey} direction={otfSortOrder} onSort={(key) => toggleColumnSort(otfSortKey, setOtfSortKey, otfSortOrder, setOtfSortOrder, key)} />
                      <SortableHeader label="Screen / Complex Name" sortKey="screenName" activeKey={otfSortKey} direction={otfSortOrder} onSort={(key) => toggleColumnSort(otfSortKey, setOtfSortKey, otfSortOrder, setOtfSortOrder, key)} />
                      <SortableHeader label="City / Town" sortKey="location" activeKey={otfSortKey} direction={otfSortOrder} onSort={(key) => toggleColumnSort(otfSortKey, setOtfSortKey, otfSortOrder, setOtfSortOrder, key)} />
                      <SortableHeader label="State" sortKey="state" activeKey={otfSortKey} direction={otfSortOrder} onSort={(key) => toggleColumnSort(otfSortKey, setOtfSortKey, otfSortOrder, setOtfSortOrder, key)} />
                      <th className="billings-page__otf-table-action-header">Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {pagedOtfTransactions.map((transaction) => {
                      const isActive =
                        activeOtfTransaction?.transactionId === transaction.transactionId;

                      return (
                        <tr
                          key={transaction.transactionId}
                          className={
                            isActive
                              ? "billings-page__otf-table-row billings-page__otf-table-row--active"
                              : "billings-page__otf-table-row"
                          }
                        >
                          <td title={transaction.complexCode || "-"}>
                            {transaction.complexCode || "-"}
                          </td>
                          <td title={transaction.screenCode || transaction.siteScope || "-"}>
                            {transaction.screenCode || transaction.siteScope || "-"}
                          </td>
                          <td title={transaction.screenName || "-"}>
                            {transaction.screenName || "-"}
                          </td>
                          <td title={transaction.location || "-"}>
                            {transaction.location || "-"}
                          </td>
                          <td title={transaction.state || "-"}>
                            {transaction.state || "-"}
                          </td>
                          <td className="billings-page__otf-table-action-cell">
                            <button
                              type="button"
                              className="billings-page__otf-open-button"
                              onClick={() => handleOpenOtfTransaction(transaction.transactionId)}
                              aria-label={`Open and edit ${transaction.screenName || transaction.billingCode || "OTF transaction"}`}
                              title="Open / Edit"
                            >
                              Open / Edit
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, marginTop: 12, flexWrap: "wrap" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                  <span className="billings-page__helper" style={{ margin: 0 }}>
                    Showing {pagedOtfTransactions.length} of {filteredOtfTransactions.length} records
                  </span>
                  <div className="billings-page__field" style={{ marginBottom: 0, minWidth: 110 }}>
                    <select
                      id="stage3-otf-rpp-bottom"
                      aria-label="Rows per page"
                      value={otfRowsPerPage}
                      onChange={(event) => setOtfRowsPerPage(Number(event.target.value))}
                    >
                      {rowsPerPageOptions.map((option) => (
                        <option key={option} value={option}>{option}</option>
                      ))}
                    </select>
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <button
                    type="button"
                    className="billings-page__secondary-button"
                    onClick={() => setOtfPage((page) => Math.max(1, page - 1))}
                    disabled={otfPage <= 1}
                  >
                    Previous
                  </button>
                  <span>Page {otfPage} of {otfPageCount}</span>
                  <button
                    type="button"
                    className="billings-page__secondary-button"
                    onClick={() => setOtfPage((page) => Math.min(otfPageCount, page + 1))}
                    disabled={otfPage >= otfPageCount}
                  >
                    Next
                  </button>
                </div>
              </div>

              {activeOtfTransaction && (
                <section className="billings-page__otf-details">
                  <div className="billings-page__billing-section-heading">
                    <h2>OTF Transaction Details</h2>
                    <span className="billings-page__billing-section-text">
                      Read-only commercial values stay preserved while invoice and RICBR references can be updated.
                    </span>
                  </div>

                  <div className="billings-page__otf-details-grid">
                    <div className="billings-page__field">
                      <label>Billing Code / Customer Code</label>
                      <input type="text" value={activeOtfTransaction.billingCode || ""} readOnly />
                    </div>

                    <div className="billings-page__field">
                      <label>Complex Code</label>
                      <input
                        type="text"
                        value={getDisplayComplexCode(activeOtfTransaction.complexCode)}
                        readOnly
                      />
                    </div>

                    <div className="billings-page__field billings-page__otf-details-grid-span">
                      <label>Included Screen Codes</label>
                      <textarea
                        rows={3}
                        value={(activeOtfTransaction.includedSiteIds || []).join(", ")}
                        readOnly
                      />
                    </div>

                    <div className="billings-page__field">
                      <label>OTF Type</label>
                      <input type="text" value={activeOtfTransaction.otfType || ""} readOnly />
                    </div>

                    <div className="billings-page__field">
                      <label>OTF Amount Before GST</label>
                      <input type="text" value={activeOtfTransaction.otfAmount || ""} readOnly />
                    </div>

                    <div className="billings-page__field">
                      <label>Invoice No.</label>
                      <input
                        type="text"
                        value={activeOtfTransaction.otfInvoiceNumber || ""}
                        onChange={(event) =>
                          updateActiveOtfTransaction("otfInvoiceNumber", event.target.value)
                        }
                        placeholder="Invoice number"
                      />
                    </div>

                    <div className="billings-page__field">
                      <label>Invoice Date</label>
                      <input
                        type="date"
                        value={activeOtfTransaction.otfInvoiceDate || ""}
                        onChange={handleValidatedOtfDateChange("otfInvoiceDate")}
                        onBlur={handleValidatedDateBlur("otfInvoiceDate")}
                      />
                    </div>

                    <div className="billings-page__field">
                      <label>RICBR / Receipt Number</label>
                      <input
                        type="text"
                        value={activeOtfTransaction.otfRicbrNumber || ""}
                        onChange={(event) =>
                          updateActiveOtfTransaction("otfRicbrNumber", event.target.value)
                        }
                        placeholder="Receipt / RICBR number"
                      />
                    </div>

                    <div className="billings-page__field">
                      <label>RICBR / Receipt Date</label>
                      <input
                        type="date"
                        value={activeOtfTransaction.otfRicbrCreatedDate || ""}
                        onChange={handleValidatedOtfDateChange("otfRicbrCreatedDate")}
                        onBlur={handleValidatedDateBlur("otfRicbrCreatedDate")}
                      />
                    </div>

                    <div className="billings-page__field">
                      <label>Received Amount</label>
                      <input
                        type="text"
                        inputMode="decimal"
                        value={activeOtfTransaction.otfReceivedAmount || ""}
                        onChange={(event) =>
                          updateActiveOtfTransaction(
                            "otfReceivedAmount",
                            sanitizeAmount(event.target.value),
                          )
                        }
                        placeholder="Received amount"
                      />
                    </div>
                  </div>

                  <div className="billings-page__otf-details-actions">
                    <button
                      type="button"
                      className="billings-page__primary-button"
                      onClick={handleSaveOtfDetails}
                    >
                      Save OTF Details
                    </button>
                  </div>
                </section>
              )}
            </>
          )}
        </article>
      )}

      {isFocBillableEditorOpen && activeFocRecord ? (
        <div
          className="billings-page__modal-backdrop"
          role="presentation"
          onClick={handleCloseFocBillableEditor}
        >
          <div
            className="billings-page__modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="foc-billable-title"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="billings-page__modal-header">
              <div>
                <p className="billings-page__modal-eyebrow">FoC Lifecycle</p>
                <h3 id="foc-billable-title">Change to Billable</h3>
              </div>
              <button
                type="button"
                className="billings-page__modal-close-button"
                onClick={handleCloseFocBillableEditor}
              >
                X
              </button>
            </div>

            <div className="billings-page__modal-grid">
              <div className="billings-page__modal-label">Screen</div>
              <div className="billings-page__modal-value">
                {activeFocRecord.screenCode || "-"} / {activeFocRecord.screenName || "-"}
              </div>

              <div className="billings-page__modal-label">Effective From</div>
              <div className="billings-page__modal-value">
                <input
                  type="date"
                  value={focBillableDraft.effectiveFrom}
                  onChange={(event) => updateFocBillableDraft("effectiveFrom", event.target.value)}
                />
              </div>

              <div className="billings-page__modal-label">OTF Applicable</div>
              <div className="billings-page__modal-value">
                <select
                  value={focBillableDraft.otfApplicable}
                  onChange={(event) => updateFocBillableDraft("otfApplicable", event.target.value)}
                >
                  <option value="No">No</option>
                  <option value="Yes">Yes</option>
                </select>
              </div>

              <div className="billings-page__modal-label">OTF Amount</div>
              <div className="billings-page__modal-value">
                <input
                  type="text"
                  inputMode="decimal"
                  value={focBillableDraft.otfAmount}
                  disabled={focBillableDraft.otfApplicable !== "Yes"}
                  onChange={(event) =>
                    updateFocBillableDraft("otfAmount", sanitizeAmount(event.target.value))
                  }
                  placeholder="Enter OTF Amount"
                />
              </div>

              <div className="billings-page__modal-label">Subscription Applicable</div>
              <div className="billings-page__modal-value">
                <select
                  value={focBillableDraft.subscriptionApplicable}
                  onChange={(event) =>
                    updateFocBillableDraft("subscriptionApplicable", event.target.value)
                  }
                >
                  <option value="No">No</option>
                  <option value="Yes">Yes</option>
                </select>
              </div>

              <div className="billings-page__modal-label">Subscription Fee</div>
              <div className="billings-page__modal-value">
                <input
                  type="text"
                  inputMode="decimal"
                  value={focBillableDraft.subscriptionFee}
                  disabled={focBillableDraft.subscriptionApplicable !== "Yes"}
                  onChange={(event) =>
                    updateFocBillableDraft(
                      "subscriptionFee",
                      sanitizeAmount(event.target.value),
                    )
                  }
                  placeholder="Enter Subscription Fee"
                />
              </div>

              <div className="billings-page__modal-label">Remarks</div>
              <div className="billings-page__modal-value">
                <textarea
                  rows={4}
                  value={focBillableDraft.remarks}
                  onChange={(event) => updateFocBillableDraft("remarks", event.target.value)}
                  placeholder="Reason for changing FoC commercial terms"
                />
              </div>
            </div>

            {focSaveMessage ? (
              <div className="billings-page__modal-notice">{focSaveMessage}</div>
            ) : null}

            <div className="billings-page__modal-footer">
              <button
                type="button"
                className="billings-page__primary-button"
                onClick={handleSaveFocBillableChange}
              >
                Save Billable Terms
              </button>
              <button
                type="button"
                className="billings-page__secondary-button"
                onClick={handleCloseFocBillableEditor}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      ) : null}

      {humanErrorCorrectionRecord && (
        <div
          className="billings-page__modal-backdrop"
          role="presentation"
          onClick={handleCloseHumanErrorCorrection}
        >
          <div
            className="billings-page__modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="human-error-correction-title"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="billings-page__modal-header">
              <div>
                <p className="billings-page__modal-eyebrow">Record Correction</p>
                <h3 id="human-error-correction-title">Request Correction</h3>
              </div>
              <button
                type="button"
                className="billings-page__modal-close-button"
                onClick={handleCloseHumanErrorCorrection}
                aria-label="Close correction request"
              >
                X
              </button>
            </div>

            <p className="billings-page__helper">
              The completed Billing Record stays unchanged as history. MoU &
              Subscription corrections move the selected screen directly to Stage 1.
            </p>

            <div className="billings-page__modal-grid">
              <div className="billings-page__modal-label">Billing Record</div>
              <div className="billings-page__modal-value">
                {humanErrorCorrectionBillingRecord?.billingRecordId || "-"}
              </div>

              <div className="billings-page__modal-label">Screen Code</div>
              <div className="billings-page__modal-value">
                {humanErrorCorrectionRecord.screenCode || "-"}
              </div>

              <div className="billings-page__modal-label">Screen Name</div>
              <div className="billings-page__modal-value">
                {humanErrorCorrectionRecord.screenName || "-"}
              </div>

              <div className="billings-page__modal-label">Correction Fields</div>
              <div className="billings-page__modal-value">
                <div style={{ display: "grid", gap: 8 }}>
                  {humanErrorCorrectionFields.map((field) => (
                    <label
                      key={field.key}
                      style={{ display: "flex", gap: 8, alignItems: "center" }}
                    >
                      <input
                        type="checkbox"
                        checked={humanErrorCorrectionFieldsSelected.includes(
                          field.key,
                        )}
                        onChange={(event) =>
                          handleToggleHumanErrorCorrectionField(
                            field.key,
                            event.target.checked,
                          )
                        }
                      />
                      {field.label}
                    </label>
                  ))}
                </div>
              </div>

              <div className="billings-page__modal-label">
                Correction Reason / Remarks
              </div>
              <div className="billings-page__modal-value">
                <textarea
                  rows={4}
                  value={humanErrorCorrectionRemarks}
                  onChange={(event) =>
                    setHumanErrorCorrectionRemarks(event.target.value)
                  }
                  placeholder="Explain the human error and what needs to be corrected"
                />
              </div>
            </div>

            {humanErrorCorrectionFieldsSelected.includes("subscriptionMode") && (
              <div className="billings-page__modal-notice">
                Subscription Mode correction requires Billing Date review in
                Stage 2 before the site returns to Stage 3.
              </div>
            )}

            <div className="billings-page__modal-footer">
              <button
                type="button"
                className="billings-page__primary-button"
                onClick={handleSubmitHumanErrorCorrection}
              >
                Send to Stage 1
              </button>
              <button
                type="button"
                className="billings-page__secondary-button"
                onClick={handleCloseHumanErrorCorrection}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {allocationValidationPopup && (
        <div
          className="billings-page__modal-backdrop"
          role="presentation"
          onClick={() => setAllocationValidationPopup(null)}
        >
          <div
            className="billings-page__modal billings-page__modal--allocation-validation"
            role="dialog"
            aria-modal="true"
            aria-labelledby="allocation-validation-title"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="billings-page__modal-header">
              <div>
                <p className="billings-page__modal-eyebrow">Price Change Validation</p>
                <h3 id="allocation-validation-title">
                  {allocationValidationPopup.title}
                </h3>
              </div>
            </div>
            <div className="billings-page__modal-notice">
              <p>{allocationValidationPopup.message}</p>
              {allocationValidationPopup.newFee && (
                <dl className="billings-page__allocation-validation-values">
                  <div><dt>New Common Subscription Fee</dt><dd>₹{allocationValidationPopup.newFee}</dd></div>
                  <div><dt>Allocation Total</dt><dd>₹{allocationValidationPopup.allocationTotal}</dd></div>
                  <div><dt>Difference</dt><dd>₹{allocationValidationPopup.difference}</dd></div>
                </dl>
              )}
              {allocationValidationPopup.differenceMessage && (
                <p>{allocationValidationPopup.differenceMessage}</p>
              )}
            </div>
            <div className="billings-page__modal-footer">
              <button
                type="button"
                className="billings-page__primary-button"
                onClick={handleGoToPriceChangeAllocation}
              >
                Go to Allocation
              </button>
              <button
                type="button"
                className="billings-page__secondary-button"
                onClick={() => setAllocationValidationPopup(null)}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {activeCommonComplexBillingRow && (
        <div
          className="billings-page__modal-backdrop"
          role="presentation"
          onClick={handleCloseCommonComplexAllocationModal}
        >
          <div
            className="billings-page__modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="common-complex-allocation-title"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="billings-page__modal-header">
              <div>
                <p className="billings-page__modal-eyebrow">
                  Site-wise Billing Allocation
                </p>
                <h3 id="common-complex-allocation-title">
                  Site-wise Billing Allocation
                </h3>
              </div>
              <button
                type="button"
                className="billings-page__modal-close-button"
                onClick={handleCloseCommonComplexAllocationModal}
                aria-label="Close allocation popup"
              >
                X
              </button>
            </div>

            <div className="billings-page__modal-grid">
              <div className="billings-page__modal-label">Billing Code / Customer Code</div>
              <div className="billings-page__modal-value">
                {normalizeValue(activeRecord?.billingCode) || "-"}
              </div>

              <div className="billings-page__modal-label">Complex Code</div>
              <div className="billings-page__modal-value">
                {getDisplayComplexCode(activeRecord?.complexCode)}
              </div>

              <div className="billings-page__modal-label">Screen Code</div>
              <div className="billings-page__modal-value">
                {activeCommonComplexBillingRow.screenCode || "-"}
              </div>

              <div className="billings-page__modal-label">Screen Name</div>
              <div className="billings-page__modal-value">
                {activeCommonComplexBillingRow.screenName || "-"}
              </div>

              <div className="billings-page__modal-label">Billing Start Date</div>
              <div className="billings-page__modal-value">
                {activeCommonComplexBillingRow.billingStartDate || "-"}
              </div>

              <div className="billings-page__modal-label">Billing Status</div>
              <div className="billings-page__modal-value">
                {activeCommonComplexBillingRow.billingStatus || "-"}
              </div>

              <div className="billings-page__modal-label">
                Common Subscription Fee
              </div>
              <div className="billings-page__modal-value">
                {commonComplexBillingSummary?.commonSubscriptionFee || "-"}
              </div>

              <div className="billings-page__modal-label">Allocated Fee</div>
              <div className="billings-page__modal-value">
                <input
                  type="text"
                  inputMode="decimal"
                  value={allocationDraftFee}
                  onChange={(event) =>
                    setAllocationDraftFee(sanitizeAmount(event.target.value))
                  }
                  readOnly={activeCommonComplexBillingRow.locked}
                  placeholder="Enter allocation"
                />
              </div>

              <div className="billings-page__modal-label">Current Billing Charge</div>
              <div className="billings-page__modal-value">
                {activeCommonComplexBillingRow.currentBillingCharge || "-"}
              </div>

              <div className="billings-page__modal-label">Created At</div>
              <div className="billings-page__modal-value">
                {formatTimestampValue(
                  activeCommonComplexBillingRow.billingAllocationCreatedAt,
                )}
              </div>

              <div className="billings-page__modal-label">Last Updated At</div>
              <div className="billings-page__modal-value">
                {formatTimestampValue(
                  activeCommonComplexBillingRow.billingAllocationUpdatedAt,
                )}
              </div>
            </div>

            {allocationSaveMessage && (
              <div className="billings-page__modal-notice">
                {allocationSaveMessage}
              </div>
            )}

            <div className="billings-page__modal-footer">
              <button
                type="button"
                className="billings-page__primary-button"
                onClick={handleSaveCommonComplexAllocation}
                disabled={activeCommonComplexBillingRow.locked}
              >
                Save Allocation
              </button>
              <button
                type="button"
                className="billings-page__secondary-button"
                onClick={handleCloseCommonComplexAllocationModal}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}

export default Billings;
