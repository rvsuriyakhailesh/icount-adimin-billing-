import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import "./App.css";
import "./LightTheme.css";
import AppLayout from "./components/layout/AppLayout";
import SiteEntry from "./pages/SiteEntry/SiteEntry";
import Installations from "./pages/Installations/Installations";
import Billings from "./pages/Billings/Billings";
import Incentive from "./pages/Incentive/Incentive";
import MouStatus from "./pages/MouStatus/MouStatus";
import Settings from "./pages/Settings/Settings";
import HistoryTracking from "./pages/HistoryTracking/HistoryTracking";
import AuditHistory from "./pages/AuditHistory/AuditHistory";
import Downloads from "./pages/Downloads/Downloads";
import { normalizeIdInput } from "./utils/idValidation";

const BILLING_API_BASE =
  import.meta.env.VITE_BILLING_API_BASE || "http://localhost:4100/api";

function normalizeValue(value) {
  return String(value || "").trim();
}


function mapOtfTypeFromBackend(value) {
  if (value === "NonRefundable") return "Non-Refundable";
  if (value === "Refundable") return "Refundable";
  return "";
}

function mapTaxModeFromBackend(value) {
  if (value === "IncludingTax") return "Including Tax";
  if (value === "ExcludingTax") return "Excluding Tax";
  if (value === "NotApplicable") return "Not Applicable";
  return "";
}

function mapSubscriptionTypeFromBackend(value) {
  const normalized = String(value ?? "").trim().toLowerCase();

  if (["vary", "variable"].includes(normalized)) return "Vary";
  if (normalized === "fixed") return "Fixed";
  return "";
}

function mapSubscriptionModeFromBackend(value) {
  if (value === "HalfYearly") return "Half-Yearly";
  if (["Monthly", "Annual"].includes(value)) return value;
  return "";
}

function mapPricingMethodFromBackend(value) {
  if (value === "sitewise") return "Site-wise";
  if (value === "allocation") return "Allocation Model";
  if (value === "common") return "Common";
  return "";
}

function decimalInputValue(value) {
  if (value === null || value === undefined || value === "") return "";
  return String(value);
}

function getLatestApprovedCommonPriceChange(site, sites = [], referenceDate = new Date()) {
  const billingCode = normalizeIdInput(site?.billingId);
  const complexCode = normalizeIdInput(site?.complexId, { allowStandaloneBlank: true });
  const targetScreenCode = normalizeIdInput(site?.siteId);
  if (!billingCode || !complexCode || site?.foc === true) return null;

  const referenceDay = new Date(referenceDate);
  referenceDay.setHours(23, 59, 59, 999);
  return (Array.isArray(sites) ? sites : [])
    .filter(
      (candidate) =>
        candidate?.foc !== true &&
        normalizeIdInput(candidate?.billingId) === billingCode &&
        normalizeIdInput(candidate?.complexId, { allowStandaloneBlank: true }) === complexCode,
    )
    .flatMap((candidate) => {
      const snapshot = candidate?.stage1Data && typeof candidate.stage1Data === "object"
        ? candidate.stage1Data
        : {};
      const inSelectedScope =
        normalizeIdInput(candidate?.siteId) === targetScreenCode ||
        (Array.isArray(snapshot.stage1GroupRows) &&
          snapshot.stage1GroupRows.some(
            (row) => normalizeIdInput(row?.screenCode) === targetScreenCode,
          ));
      if (!inSelectedScope) return [];

      const stage3 = candidate?.stage3Data && typeof candidate.stage3Data === "object"
        ? candidate.stage3Data
        : {};
      return (Array.isArray(stage3.priceChangeHistory) ? stage3.priceChangeHistory : [])
        .filter((entry) => {
          const effectiveDate = new Date(`${String(entry?.effectiveDate || "")}T00:00:00`);
          return (
            normalizeValue(entry?.scope).toLowerCase() === "common complex" &&
            normalizeValue(entry?.approvalStatus).toLowerCase() === "not required" &&
            Number.isFinite(effectiveDate.getTime()) &&
            effectiveDate <= referenceDay &&
            normalizeValue(entry?.newFee) !== ""
          );
        })
        .map((entry) => ({
          ...entry,
          _effectiveAt: new Date(`${entry.effectiveDate}T00:00:00`).getTime(),
          _approvedAt: Date.parse(entry.validatedAt || entry.savedAndSubmittedAt || "") || 0,
        }));
    })
    .sort(
      (left, right) =>
        right._effectiveAt - left._effectiveAt || right._approvedAt - left._approvedAt,
    )[0] || null;
}

function getLatestCommonCommercialTerm(site, sites = []) {
  const billingCode = normalizeIdInput(site?.billingId);
  const complexCode = normalizeIdInput(site?.complexId, { allowStandaloneBlank: true });
  const targetScreenCode = normalizeIdInput(site?.siteId);
  const candidates = (Array.isArray(sites) ? sites : [])
    .filter(
      (candidate) =>
        candidate?.foc !== true &&
        normalizeIdInput(candidate?.billingId) === billingCode &&
        normalizeIdInput(candidate?.complexId, { allowStandaloneBlank: true }) === complexCode,
    )
    .filter((candidate) => {
      const snapshot = candidate?.stage1Data && typeof candidate.stage1Data === "object"
        ? candidate.stage1Data
        : {};
      return (
        normalizeIdInput(candidate?.siteId) === targetScreenCode ||
        (Array.isArray(snapshot.stage1GroupRows) &&
          snapshot.stage1GroupRows.some(
            (row) => normalizeIdInput(row?.screenCode) === targetScreenCode,
          ))
      );
    })
    .flatMap((candidate) => (Array.isArray(candidate?.commercialTerms) ? candidate.commercialTerms : []))
    .filter((term) => term?.isCurrent !== false && term?.pricingMode === "common")
    .sort(
      (left, right) =>
        (Date.parse(right.updatedAt || right.createdAt || "") || 0) -
          (Date.parse(left.updatedAt || left.createdAt || "") || 0) ||
        (Number(right.version) || 0) - (Number(left.version) || 0),
    );
  const baseTerm = candidates[0] || (Array.isArray(site?.commercialTerms) ? site.commercialTerms[0] : null);
  const priceChange = getLatestApprovedCommonPriceChange(site, sites);
  return priceChange
    ? {
        ...baseTerm,
        subscriptionFee: priceChange.newFee,
        subscriptionMode: priceChange.newMode || baseTerm?.subscriptionMode,
        _sourceUpdatedAt: priceChange._approvedAt || priceChange._effectiveAt,
      }
    : baseTerm;
}

function getStage1RecordIdentity(record) {
  return {
    backendSiteId: String(record?.backendSiteId || "").trim(),
    screenCode: normalizeIdInput(record?.screenCode),
  };
}

function upsertStage1Records(currentRecords, incomingRecords) {
  const incoming = Array.isArray(incomingRecords) ? incomingRecords : [];
  const incomingIdentities = incoming.map(getStage1RecordIdentity);
  const retained = currentRecords.filter((record) => {
    const identity = getStage1RecordIdentity(record);

    return !incomingIdentities.some(
      (incomingIdentity) =>
        (identity.backendSiteId &&
          identity.backendSiteId === incomingIdentity.backendSiteId) ||
        (identity.screenCode &&
          identity.screenCode === incomingIdentity.screenCode),
    );
  });
  const uniqueIncoming = [];

  incoming.forEach((record) => {
    const identity = getStage1RecordIdentity(record);
    const duplicate = uniqueIncoming.some((existingRecord) => {
      const existingIdentity = getStage1RecordIdentity(existingRecord);
      return (
        (identity.backendSiteId &&
          identity.backendSiteId === existingIdentity.backendSiteId) ||
        (identity.screenCode && identity.screenCode === existingIdentity.screenCode)
      );
    });

    if (!duplicate) {
      uniqueIncoming.push(record);
    }
  });

  return [...retained, ...uniqueIncoming];
}

async function fetchBackendSites() {
  const response = await fetch(`${BILLING_API_BASE}/sites`);
  const body = await response.json().catch(() => ({}));

  if (!response.ok) {
    throw new Error(
      body?.message || body?.error || `Unable to load Stage 2 sites (${response.status}).`,
    );
  }

  return Array.isArray(body?.data) ? body.data : [];
}

function dedupeInstallationExpenses(expenses, screenCode = "") {
  const seen = new Set();

  return (Array.isArray(expenses) ? expenses : []).filter((expense) => {
    const expenseId = String(expense?.expenseId || "").trim() || [
      String(expense?.screenCode || screenCode).trim().toUpperCase(),
      String(expense?.expenseType || "").trim().toLowerCase(),
      String(expense?.expenseDate || "").trim(),
      Number(expense?.amount || 0),
      String(expense?.comments || "").trim(),
      String(expense?.createdAt || expense?.updatedAt || "").trim(),
    ].join("::");

    if (seen.has(expenseId)) return false;
    seen.add(expenseId);
    return true;
  });
}

function buildStage2RecordFromBackendSite(site, sites = []) {
  const stage1Snapshot =
    site?.stage1Data && typeof site.stage1Data === "object"
      ? site.stage1Data
      : {};
  const stage2Snapshot =
    site?.stage2Data && typeof site.stage2Data === "object"
      ? site.stage2Data
      : {};
  const snapshot = { ...stage1Snapshot, ...stage2Snapshot };
  const persistedOtfSnapshot =
    site?.stage3Data && typeof site.stage3Data === "object"
      ? site.stage3Data
      : {};
  const currentCommercial = getLatestCommonCommercialTerm(site, sites);
  const snapshotHas = (key) =>
    Object.prototype.hasOwnProperty.call(snapshot, key) &&
    snapshot[key] !== null &&
    snapshot[key] !== undefined &&
    snapshot[key] !== "";
  const otfValue = (key) =>
    snapshotHas(key) ? snapshot[key] : persistedOtfSnapshot[key] || "";

  const installationExpenses = dedupeInstallationExpenses(
    snapshot.installationExpenses,
    site.siteId,
  );

  return {
    ...snapshot,
    backendSiteId: site.id,
    recordId: snapshot.recordId || site.id,
    billingCode: site.billingId || "",
    complexCode: site.complexId || "",
    screenCode: site.siteId || "",
    screenName: site.screenName || "",
    billingName: snapshot.billingName || site.screenName || "",
    location: site.location || "",
    state: site.state || "",
    siteType:
      snapshot.siteType || (site.siteType === "Single" ? "Standalone" : "Complex"),
    foc: site.foc === true,

    // Carry Stage 1 commercial data into Stage 2 hydration. This is required
    // when an FoC screen is returned to Stage 1 and 1B must use a non-FoC
    // sibling in the same complex as its commercial reference.
    otfApplicable: snapshotHas("otfApplicable")
      ? snapshot.otfApplicable
      : currentCommercial?.otfApplicable || "No",
    otfType: snapshotHas("otfType")
      ? snapshot.otfType
      : mapOtfTypeFromBackend(currentCommercial?.otfType),
    otfTaxMode: snapshotHas("otfTaxMode")
      ? snapshot.otfTaxMode
      : mapTaxModeFromBackend(currentCommercial?.otfTaxMode),
    otfAmount: snapshotHas("otfAmount")
      ? snapshot.otfAmount
      : decimalInputValue(currentCommercial?.otfAmount),
    otfInvoiceNumber: otfValue("otfInvoiceNumber"),
    otfInvoiceDate: otfValue("otfInvoiceDate"),
    otfRicbrNumber: otfValue("otfRicbrNumber"),
    otfRicbrCreatedDate: otfValue("otfRicbrCreatedDate"),
    otfReceiptNumber: otfValue("otfReceiptNumber"),
    otfReceiptDate: otfValue("otfReceiptDate"),
    otfReceivedAmount: otfValue("otfReceivedAmount"),
    otfPaymentReceivedAmount: otfValue("otfPaymentReceivedAmount"),
    otfPaymentStatus: otfValue("otfPaymentStatus"),
    otfPaymentReceivedDate: otfValue("otfPaymentReceivedDate"),
    otfPaymentReference: otfValue("otfPaymentReference"),
    otfIncludedScreenCodes: Array.isArray(persistedOtfSnapshot.otfIncludedScreenCodes)
      ? persistedOtfSnapshot.otfIncludedScreenCodes
      : [],
    otfTransactionGroupKey: otfValue("otfTransactionGroupKey"),
    otfDetailsUpdatedAt: otfValue("otfDetailsUpdatedAt"),
    subscriptionType: snapshotHas("subscriptionType")
      ? snapshot.subscriptionType
      : mapSubscriptionTypeFromBackend(currentCommercial?.subscriptionType),
    subscriptionMode: snapshotHas("subscriptionMode")
      ? snapshot.subscriptionMode
      : mapSubscriptionModeFromBackend(currentCommercial?.subscriptionMode),
    subscriptionFee: snapshotHas("subscriptionFee")
      ? snapshot.subscriptionFee
      : decimalInputValue(currentCommercial?.subscriptionFee),
    pricingMethod: snapshotHas("pricingMethod")
      ? snapshot.pricingMethod
      : mapPricingMethodFromBackend(currentCommercial?.pricingMode),
    pricingGroups: Array.isArray(snapshot.pricingGroups)
      ? snapshot.pricingGroups.map((group) => ({ ...group }))
      : [],
    allocationRows: Array.isArray(snapshot.allocationRows)
      ? snapshot.allocationRows.map((row) => ({ ...row }))
      : [],
    incentiveApplicable:
      snapshot.incentiveApplicable ?? site.incentiveApplicable ?? false,
    salesEmployeeName:
      snapshot.salesEmployeeName || site.incentiveEmployeeName || "",
    companyId: snapshot.companyId || site.incentiveEmployeeNumber || "",

    processingStatus: site.processingStatus || snapshot.processingStatus || "Stage 2",
    readinessStatus: site.readinessStatus || snapshot.readinessStatus || "Moved to Stage 2",
    selected: false,
    currentStageStatus: snapshot.currentStageStatus || "",
    installationStatus: snapshot.installationStatus || "",
    dateOfDispatch: snapshot.dateOfDispatch || "",
    installationDate: snapshot.installationDate || "",
    liveDate: snapshot.liveDate || "",
    trialPeriod: snapshot.trialPeriod || "",
    trialPeriodExtension: snapshot.trialPeriodExtension || "",
    totalTrialPeriodExtension: snapshot.totalTrialPeriodExtension || "",
    trialExtension: snapshot.trialExtension || "",
    blockerReason: snapshot.blockerReason || "",
    remarks: snapshot.remarks || "",
    extensionRemarks: snapshot.extensionRemarks || "",
    billingStartDate: snapshot.billingStartDate || "",
    installationDetailsCompleted: snapshot.installationDetailsCompleted === true,
    saved2B: snapshot.saved2B === true,
    saved2BAt: snapshot.saved2BAt || "",
    installationExpenses,
  };
}

function buildStage3RecordFromBackendSite(site, sites = []) {
  const baseRecord = buildStage2RecordFromBackendSite(site, sites);
  const stage3Snapshot =
    site?.stage3Data && typeof site.stage3Data === "object"
      ? site.stage3Data
      : {};
  const currentCommercial =
    Array.isArray(site?.commercialTerms) && site.commercialTerms.length > 0
      ? site.commercialTerms[0]
      : null;
  const pricingRows = Array.isArray(currentCommercial?.pricingRows)
    ? currentCommercial.pricingRows
    : [];
  const currentScreenCode = normalizeIdInput(site?.siteId);
  const currentPricingRow =
    pricingRows.find(
      (row) => normalizeIdInput(row?.screenCode) === currentScreenCode,
    ) || pricingRows[0] || null;

  const persistedAllocationRows = Array.isArray(stage3Snapshot?.allocationRows)
    ? stage3Snapshot.allocationRows.map((row) => ({ ...row }))
    : Array.isArray(baseRecord?.allocationRows)
      ? baseRecord.allocationRows.map((row) => ({ ...row }))
      : [];

  if (currentPricingRow && currentCommercial?.pricingMode === "allocation") {
    const targetIndex = persistedAllocationRows.findIndex(
      (row) => normalizeIdInput(row?.screenCode) === currentScreenCode,
    );
    const backendRow = {
      ...(targetIndex >= 0 ? persistedAllocationRows[targetIndex] : {}),
      screenCode: site.siteId,
      screenName: site.screenName,
      deviceCount:
        currentPricingRow.deviceCount ??
        persistedAllocationRows[targetIndex]?.deviceCount ??
        persistedAllocationRows[targetIndex]?.count ??
        null,
      count:
        currentPricingRow.deviceCount ??
        persistedAllocationRows[targetIndex]?.count ??
        persistedAllocationRows[targetIndex]?.deviceCount ??
        null,
      planMode:
        currentPricingRow.planMode ||
        persistedAllocationRows[targetIndex]?.planMode ||
        "",
      planFee:
        decimalInputValue(currentPricingRow.planFee) ||
        persistedAllocationRows[targetIndex]?.planFee ||
        "",
    };

    if (targetIndex >= 0) persistedAllocationRows[targetIndex] = backendRow;
    else persistedAllocationRows.push(backendRow);
  }

  return {
    ...baseRecord,
    ...stage3Snapshot,
    backendSiteId: site.id,
    recordId: stage3Snapshot.recordId || baseRecord.recordId || site.id,
    billingCode: site.billingId || "",
    complexCode: site.complexId || "",
    screenCode: site.siteId || "",
    screenName: site.screenName || "",
    location: site.location || "",
    state: site.state || "",
    processingStatus: site.processingStatus || "",
    readinessStatus: site.readinessStatus || "",
    allocationRows: persistedAllocationRows,
    installationExpenses: dedupeInstallationExpenses(
      stage3Snapshot.installationExpenses || baseRecord.installationExpenses,
      site.siteId,
    ),
  };
}

function isCurrentStage3Record(record) {
  const processingStatus = normalizeValue(record?.processingStatus).toLowerCase();
  const readinessStatus = normalizeValue(record?.readinessStatus).toLowerCase();
  return (
    processingStatus === "stage 3" ||
    readinessStatus === "moved to stage 3"
  );
}

function buildStage1RecordFromBackendSite(site, sites = []) {
  const baseRecord = buildStage2RecordFromBackendSite(site, sites);
  const stage1Snapshot =
    site?.stage1Data && typeof site.stage1Data === "object"
      ? site.stage1Data
      : {};

  return {
    ...baseRecord,
    ...stage1Snapshot,
    backendSiteId: site.id,
    recordId: stage1Snapshot.recordId || baseRecord.recordId || site.id,
    billingCode: site.billingId || "",
    complexCode: site.complexId || "",
    screenCode: site.siteId || "",
    screenName: site.screenName || "",
    location: site.location || "",
    state: site.state || "",
    foc: site.foc === true,
    processingStatus: site.processingStatus || "Stage 1",
    readinessStatus: site.readinessStatus || "Ready",
    selected: false,
  };
}

function getPersistedCommonComplexBillingAllocation(record) {
  const snapshot = record?.commonComplexBillingAllocation;
  if (!snapshot || typeof snapshot !== "object" || Array.isArray(snapshot)) {
    return null;
  }

  const groupKey = String(
    snapshot.groupKey ||
      record?.stage1GroupId ||
      record?.complexCode ||
      record?.billingCode ||
      "",
  ).trim();
  const allocations =
    snapshot.allocations &&
    typeof snapshot.allocations === "object" &&
    !Array.isArray(snapshot.allocations)
      ? snapshot.allocations
      : {};

  return groupKey && Object.keys(allocations).length > 0
    ? { groupKey, allocations }
    : null;
}

function deriveBillingTreatment(record) {
  const existingBillingTreatment = String(record.billingTreatment || "").trim();
  if (existingBillingTreatment) {
    return existingBillingTreatment;
  }

  const billingStartDate = String(record.billingStartDate || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(billingStartDate)) {
    return "";
  }

  const day = Number(billingStartDate.slice(8, 10));
  if (!Number.isFinite(day)) {
    return "";
  }

  return day === 1 ? "Full Month" : "Pro-rata";
}

function normalizeDateValue(value) {
  return String(value || "").trim();
}

function parseLocalDateValue(value) {
  const normalizedValue = normalizeDateValue(value);

  if (!/^\d{4}-\d{2}-\d{2}$/.test(normalizedValue)) {
    return null;
  }

  const [year, month, day] = normalizedValue.split("-").map(Number);
  const parsed = new Date(year, month - 1, day);

  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function formatLocalDateValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

function addCalendarMonths(date, months) {
  const year = date.getFullYear();
  const month = date.getMonth() + months;
  const lastDayOfTargetMonth = new Date(year, month + 1, 0).getDate();
  const day = Math.min(date.getDate(), lastDayOfTargetMonth);

  return new Date(year, month, day);
}

function getBillingPeriodFrequency(record) {
  const normalizedMode = normalizeDateValue(record.subscriptionMode)
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
    normalizedMode === "semi annual"
  ) {
    return "half-year";
  }

  if (normalizedMode === "quarterly" || normalizedMode === "quarter") {
    return "quarter";
  }

  return "month";
}

function calculateBillingPeriodTo(fromValue, record) {
  const fromDate = parseLocalDateValue(fromValue);

  if (!fromDate) {
    return "";
  }

  const frequency = getBillingPeriodFrequency(record);

  if (frequency === "year") {
    const nextDate = addCalendarMonths(fromDate, 12);
    nextDate.setDate(nextDate.getDate() - 1);
    return formatLocalDateValue(nextDate);
  }

  if (frequency === "half-year") {
    const nextDate = addCalendarMonths(fromDate, 6);
    nextDate.setDate(nextDate.getDate() - 1);
    return formatLocalDateValue(nextDate);
  }

  if (frequency === "quarter") {
    const nextDate = addCalendarMonths(fromDate, 3);
    nextDate.setDate(nextDate.getDate() - 1);
    return formatLocalDateValue(nextDate);
  }

  const lastDayOfMonth = new Date(
    fromDate.getFullYear(),
    fromDate.getMonth() + 1,
    0,
  );
  return formatLocalDateValue(lastDayOfMonth);
}

function resolveStage3ComplexId(record) {
  const directComplexId = normalizeIdInput(record?.complexCode, {
    allowStandaloneBlank: true,
  });

  if (directComplexId) {
    return directComplexId;
  }

  const screenCode = normalizeIdInput(record?.screenCode);
  const groupRows = Array.isArray(record?.stage1GroupRows)
    ? record.stage1GroupRows
    : [];

  const matchingRow = groupRows.find(
    (row) => normalizeIdInput(row?.screenCode) === screenCode,
  );

  const matchedComplexId = normalizeIdInput(matchingRow?.complexCode, {
    allowStandaloneBlank: true,
  });

  if (matchedComplexId) {
    return matchedComplexId;
  }

  const firstComplexId = groupRows
    .map((row) =>
      normalizeIdInput(row?.complexCode, { allowStandaloneBlank: true }),
    )
    .find(Boolean);

  return firstComplexId || "";
}

function getStage3BillingPeriodFields(record) {
  const billingStartDate = normalizeDateValue(record.billingStartDate);

  if (!billingStartDate) {
    return {
      billingPeriodFrom: "",
      billingPeriodTo: "",
    };
  }

  return {
    billingPeriodFrom: billingStartDate,
    billingPeriodTo: calculateBillingPeriodTo(billingStartDate, record),
  };
}

function createStage3Record(record) {
  const billingTreatment = deriveBillingTreatment(record);
  const billingPeriodFields = getStage3BillingPeriodFields(record);
  const resolvedComplexId = resolveStage3ComplexId(record);

  return {
    ...record,
    billingCode: record.billingCode,
    complexCode: resolvedComplexId,
    screenCode: record.screenCode,
    screenName: record.screenName,
    location: record.location,
    state: record.state,
    billingStartDate: record.billingStartDate,
    stage3BillingStartDateSnapshot: normalizeDateValue(record.billingStartDate),
    subscriptionType: record.subscriptionType,
    subscriptionMode: record.subscriptionMode,
    subscriptionFee: record.subscriptionFee,
    otfApplicable: record.otfApplicable,
    otfType: record.otfType,
    otfAmount: record.otfAmount,
    currentStageStatus:
      record.currentStageStatus || record.installationStatus || "In Progress",
    billingPeriodFrom: billingPeriodFields.billingPeriodFrom,
    billingPeriodTo: billingPeriodFields.billingPeriodTo,
    invoiceNumber: record.invoiceNumber || "",
    invoiceDate: record.invoiceDate || "",
    invoiceAmount: record.invoiceAmount || "",
    paymentStatus: record.paymentStatus || "Not Invoiced",
    paymentReceivedDate: record.paymentReceivedDate || "",
    billingRemarks: record.billingRemarks || "",
    billingTreatment: billingTreatment || record.billingTreatment || "",
    billingVerificationStatus:
      record.billingVerificationStatus || "First Billing Pending Approval",
    subscriptionInvoiceNumber: record.subscriptionInvoiceNumber || "",
    otfInvoiceNumber: record.otfInvoiceNumber || "",
    billingVerifiedAt: record.billingVerifiedAt || "",
    closureStatus: record.closureStatus || "Open",
    closureSelected: false,
  };
}

function mergeStage3RecordBySiteId(existingRecord, incomingRecord) {
  const nextRecord = {
    ...existingRecord,
    ...incomingRecord,
  };

  nextRecord.recordId = existingRecord.recordId || incomingRecord.recordId;
  nextRecord.selected = false;
  nextRecord.complexCode =
    resolveStage3ComplexId(incomingRecord) ||
    resolveStage3ComplexId(existingRecord);
  nextRecord.billingVerificationStatus =
    incomingRecord.billingVerificationStatus ||
    "First Billing Pending Approval";
  nextRecord.billingVerifiedAt = incomingRecord.billingVerifiedAt || "";
  nextRecord.closureStatus = "Open";
  nextRecord.closureSelected = false;

  const billingPeriodFields = getStage3BillingPeriodFields(incomingRecord);

  nextRecord.billingPeriodFrom = billingPeriodFields.billingPeriodFrom;
  nextRecord.billingPeriodTo = billingPeriodFields.billingPeriodTo;

  return nextRecord;
}

function mergeStage2RecordBySiteId(existingRecord, incomingRecord) {
  const nextRecord = {
    ...existingRecord,
    ...incomingRecord,
  };

  nextRecord.recordId = existingRecord.recordId || incomingRecord.recordId;
  nextRecord.selected = false;
  nextRecord.stage1CorrectionReturn = Boolean(
    incomingRecord.stage1CorrectionReturn || existingRecord.stage1CorrectionReturn,
  );

  const lifecycleFieldPreserves = [
    "currentStageStatus",
    "installationStatus",
    "dateOfDispatch",
    "installationDate",
    "liveDate",
    "trialPeriod",
    "trialPeriodExtension",
    "totalTrialPeriodExtension",
    "trialExtension",
    "billingStartDate",
    "blockerReason",
    "remarks",
    "installationDetailsCompleted",
    "installationCompletedAt",
    "installationDetailsSnapshot",
    "verifiedBy",
  ];

  lifecycleFieldPreserves.forEach((field) => {
    if (existingRecord[field] !== undefined) {
      nextRecord[field] = existingRecord[field];
    }
  });

  return nextRecord;
}

const stage2CorrectionFieldKeys = [
  "billingCode",
  "complexCode",
  "screenCode",
  "screenName",
  "location",
  "state",
  "siteType",
  "processingStatus",
  "mouStatus",
  "mouSentDate",
  "mouReceivedDate",
  "mouStartDate",
  "mouEndDate",
  "otfApplicable",
  "otfType",
  "otfAmount",
  "subscriptionType",
  "subscriptionMode",
  "subscriptionFee",
  "extensionRequired",
  "extensionRenewal",
  "newMouStartDate",
  "newMouEndDate",
  "extensionRemarks",
  "incentiveBeneficiary",
  "salesEmployeeName",
  "employeeName",
  "companyId",
  "pricingMethod",
  "pricingGroups",
  "agreementType",
  "addendumEffectiveFrom",
  "addendumRemarks",
  "previousAgreementSnapshot",
  "requiresBillingDateReview",
  "correctionType",
  "correctionRoute",
  "requestedAt",
  "stage1GroupId",
  "stage1GroupRows",
  "savedAt",
  "stage1CorrectionReturn",
];

function mergeCorrectedStage1FieldsIntoStage2Record(existingRecord, incomingRecord) {
  const nextRecord = { ...existingRecord };

  stage2CorrectionFieldKeys.forEach((field) => {
    if (incomingRecord[field] !== undefined) {
      nextRecord[field] = incomingRecord[field];
    }
  });

  nextRecord.recordId = existingRecord.recordId;
  nextRecord.selected = false;
  nextRecord.stage1CorrectionReturn = false;

  nextRecord.currentStageStatus = existingRecord.currentStageStatus;
  nextRecord.installationStatus = existingRecord.installationStatus;
  nextRecord.dateOfDispatch = existingRecord.dateOfDispatch;
  nextRecord.installationDate = existingRecord.installationDate;
  nextRecord.liveDate = existingRecord.liveDate;
  nextRecord.trialPeriod = existingRecord.trialPeriod;
  nextRecord.trialPeriodExtension = existingRecord.trialPeriodExtension;
  nextRecord.totalTrialPeriodExtension = existingRecord.totalTrialPeriodExtension;
  nextRecord.trialExtension = existingRecord.trialExtension;
  nextRecord.blockerReason = existingRecord.blockerReason;
  nextRecord.remarks = existingRecord.remarks;
  nextRecord.extensionRemarks = existingRecord.extensionRemarks;
  nextRecord.billingStartDate = existingRecord.billingStartDate;
  nextRecord.verifiedBy = existingRecord.verifiedBy;
  nextRecord.installationDetailsCompleted = existingRecord.installationDetailsCompleted;
  nextRecord.installationCompletedAt = existingRecord.installationCompletedAt;
  nextRecord.installationDetailsSnapshot = existingRecord.installationDetailsSnapshot;

  return nextRecord;
}

function App() {
  const [activePage, setActivePage] = useState("site-entry");
  const [stage1ReadinessRecords, setStage1ReadinessRecords] = useState([]);
  const [stage2Records, setStage2Records] = useState([]);
  const [stage3Records, setStage3Records] = useState([]);
  const [canonicalSites, setCanonicalSites] = useState([]);
  const [billingRecords, setBillingRecords] = useState([]);
  const [commonComplexAllocations, setCommonComplexAllocations] = useState({});
  const [incentiveStates, setIncentiveStates] = useState({});
  const [erpMappings, setErpMappings] = useState([
    {
      mappingId: "billing-inc-default",
      product: "Billing INC",
      glAccount: "",
      costCenter: "",
      status: "Active",
      additionalFields: [],
      updatedAt: "",
    },
  ]);
  const [stage1CorrectionRecord, setStage1CorrectionRecord] = useState(null);
  const [stage1CorrectionScreenCode, setStage1CorrectionScreenCode] = useState("");
  const [stage1CorrectionSnapshot, setStage1CorrectionSnapshot] = useState(null);
  const [stage2FocusRecordId, setStage2FocusRecordId] = useState("");
  const stage2ReentrySnapshotsRef = useRef(new Map());
  const hydratedCommonAllocationGroupsRef = useRef(new Set());
  const currentStage3Records = useMemo(
    () => stage3Records.filter(isCurrentStage3Record),
    [stage3Records],
  );
  const canonicalSiteRecords = useMemo(
    () =>
      canonicalSites
        .map((site) => buildStage3RecordFromBackendSite(site, canonicalSites))
        .filter((record) => normalizeIdInput(record.screenCode)),
    [canonicalSites],
  );

  useEffect(() => {
    let cancelled = false;

    async function restoreStage2Records() {
      try {
        const sites = await fetchBackendSites();
        if (cancelled) return;
        setCanonicalSites(sites);

        const backendStage1Records = sites
          .filter((site) => {
            const processingStatus = String(site?.processingStatus || "")
              .trim()
              .toLowerCase();
            const readinessStatus = String(site?.readinessStatus || "")
              .trim()
              .toLowerCase();

            return (
              processingStatus === "stage 1" ||
              readinessStatus === "ready"
            );
          })
          .map((site) => buildStage1RecordFromBackendSite(site, sites))
          .filter((record) => normalizeIdInput(record.screenCode));

        setStage1ReadinessRecords((currentRecords) =>
          upsertStage1Records(currentRecords, backendStage1Records),
        );

        const backendStage2Records = sites
          .filter((site) => {
            const processingStatus = String(site?.processingStatus || "")
              .trim()
              .toLowerCase();
            const readinessStatus = String(site?.readinessStatus || "")
              .trim()
              .toLowerCase();

            return (
              processingStatus === "stage 2" ||
              readinessStatus === "moved to stage 2"
            );
          })
          .map((site) => buildStage2RecordFromBackendSite(site, sites))
          .filter((record) => normalizeIdInput(record.screenCode));

        setStage2Records((currentRecords) => {
          const merged = [...currentRecords];

          backendStage2Records.forEach((incomingRecord) => {
            const normalizedScreenCode = normalizeIdInput(incomingRecord.screenCode);
            const existingIndex = merged.findIndex(
              (record) =>
                (record.backendSiteId &&
                  record.backendSiteId === incomingRecord.backendSiteId) ||
                (!record.backendSiteId &&
                  normalizeIdInput(record.screenCode) === normalizedScreenCode),
            );

            if (existingIndex >= 0) {
              // Keep any live Stage 2 edits already present in memory. Backend
              // restoration is primarily for page refresh / fresh app load.
              merged[existingIndex] = {
                ...incomingRecord,
                ...merged[existingIndex],
                backendSiteId:
                  merged[existingIndex].backendSiteId || incomingRecord.backendSiteId,
                recordId:
                  merged[existingIndex].recordId || incomingRecord.recordId,
                selected: false,
              };
              return;
            }

            merged.push(incomingRecord);
          });

          return merged;
        });
      } catch (error) {
        if (!cancelled) {
          console.error("Unable to restore Stage 2 records from backend:", error);
        }
      }
    }

    restoreStage2Records();

    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    const persistedGroups = [];

    stage3Records.forEach((record) => {
      const persisted = getPersistedCommonComplexBillingAllocation(record);
      if (
        persisted &&
        !hydratedCommonAllocationGroupsRef.current.has(persisted.groupKey)
      ) {
        persistedGroups.push(persisted);
      }
    });

    if (persistedGroups.length === 0) return;

    setCommonComplexAllocations((currentAllocations) => {
      const nextAllocations = { ...currentAllocations };
      persistedGroups.forEach(({ groupKey, allocations }) => {
        nextAllocations[groupKey] = {
          ...(nextAllocations[groupKey] || {}),
          ...allocations,
        };
        hydratedCommonAllocationGroupsRef.current.add(groupKey);
      });
      return nextAllocations;
    });
  }, [stage3Records]);

  useEffect(() => {
    let cancelled = false;

    async function restoreStage3Records() {
      try {
        const sites = await fetchBackendSites();
        if (cancelled) return;
        setCanonicalSites(sites);

        // Common commercial terms are the source of truth for unbilled work.
        // Stage snapshots and saved allocation totals are historical workflow
        // data and must not select an older common fee during hydration.
        const commonStage1Configurations = new Map();
        sites.forEach((site) => {
          const currentCommercial = getLatestCommonCommercialTerm(site, sites);
          const billingCode = normalizeIdInput(site?.billingId);
          const complexCode = normalizeIdInput(site?.complexId, {
            allowStandaloneBlank: true,
          });
          if (
            currentCommercial?.pricingMode !== "common" ||
            !billingCode ||
            !complexCode ||
            site?.foc === true
          ) {
            return;
          }

          const groupKey = `${billingCode}|${complexCode}`;
          const updatedAt =
            currentCommercial._sourceUpdatedAt ||
            Date.parse(currentCommercial.updatedAt || currentCommercial.createdAt || "") ||
            0;
          const configuration = {
            updatedAt,
            version: Number(currentCommercial.version) || 0,
            values: {
              subscriptionType: mapSubscriptionTypeFromBackend(
                currentCommercial.subscriptionType,
              ),
              subscriptionMode: mapSubscriptionModeFromBackend(
                currentCommercial.subscriptionMode,
              ),
              subscriptionFee: decimalInputValue(currentCommercial.subscriptionFee),
              otfApplicable: currentCommercial.otfApplicable || "No",
              otfType: mapOtfTypeFromBackend(currentCommercial.otfType),
              otfAmount: decimalInputValue(currentCommercial.otfAmount),
              otfTaxMode: mapTaxModeFromBackend(currentCommercial.otfTaxMode),
              pricingMethod: "Common",
            },
          };
          const snapshot =
            site?.stage1Data && typeof site.stage1Data === "object"
              ? site.stage1Data
              : {};
          const scopedScreenCodes = [
            normalizeIdInput(site?.siteId),
            ...(Array.isArray(snapshot.stage1GroupRows)
              ? snapshot.stage1GroupRows.map((row) => normalizeIdInput(row?.screenCode))
              : []),
          ].filter(Boolean);
          scopedScreenCodes.forEach((screenCode) => {
            const scopedKey = `${groupKey}|${screenCode}`;
            const scopedConfigurations =
              commonStage1Configurations.get(scopedKey) || [];
            scopedConfigurations.push(configuration);
            commonStage1Configurations.set(scopedKey, scopedConfigurations);
          });
        });

        const backendStage3Records = sites
          .filter((site) => {
            const processingStatus = String(site?.processingStatus || "")
              .trim()
              .toLowerCase();
            const readinessStatus = String(site?.readinessStatus || "")
              .trim()
              .toLowerCase();

            return (
              processingStatus === "stage 3" ||
              readinessStatus === "moved to stage 3"
            );
          })
          .map((site) => {
            const record = buildStage3RecordFromBackendSite(site, sites);
            const groupKey = `${normalizeIdInput(site?.billingId)}|${normalizeIdInput(
              site?.complexId,
              { allowStandaloneBlank: true },
            )}`;
            const scopedGroupKey = `${groupKey}|${normalizeIdInput(site?.siteId)}`;
            const authoritativeConfiguration = [
              ...(commonStage1Configurations.get(scopedGroupKey) || []),
            ].sort(
              (left, right) =>
                right.updatedAt - left.updatedAt || right.version - left.version,
            )[0]?.values;

            return createStage3Record(
              authoritativeConfiguration
                ? { ...record, ...authoritativeConfiguration }
                : record,
            );
          })
          .filter((record) => normalizeIdInput(record.screenCode));

        const persistedFirstTimeBillingRecords = backendStage3Records
          .map((record) => record?.firstTimeBillingRecord)
          .filter(
            (record) =>
              record &&
              typeof record === "object" &&
              !Array.isArray(record) &&
              String(record.billingRecordId || "").trim(),
          );

        setBillingRecords((currentRecords) => {
          const merged = [...currentRecords];

          persistedFirstTimeBillingRecords.forEach((incomingRecord) => {
            const billingRecordId = String(
              incomingRecord.billingRecordId || "",
            ).trim();
            const existingIndex = merged.findIndex(
              (record) =>
                String(record?.billingRecordId || "").trim() ===
                billingRecordId,
            );

            if (existingIndex >= 0) {
              merged[existingIndex] = {
                ...incomingRecord,
                ...merged[existingIndex],
                billingRecordId,
              };
              return;
            }

            merged.push(incomingRecord);
          });

          return merged;
        });

        const backendStage3ScreenCodes = new Set(
          backendStage3Records.map((record) => normalizeIdInput(record.screenCode)),
        );
        const backendStage3SiteIds = new Set(
          backendStage3Records.map((record) => record.backendSiteId).filter(Boolean),
        );
        const backendScreenCodes = new Set(
          sites.map((site) => normalizeIdInput(site?.siteId)).filter(Boolean),
        );
        setStage3Records((currentRecords) => {
          // Backend top-level stage fields are authoritative. Remove stale
          // local projections for known sites that are no longer in Stage 3;
          // nested stage3Data remains available from the API for history.
          const merged = currentRecords.filter((record) => {
            const screenCode = normalizeIdInput(record?.screenCode);
            const backendSiteId = record?.backendSiteId;
            return (
              !screenCode ||
              (backendSiteId
                ? backendStage3SiteIds.has(backendSiteId)
                : !backendScreenCodes.has(screenCode) ||
                  backendStage3ScreenCodes.has(screenCode))
            );
          });

          backendStage3Records.forEach((incomingRecord) => {
            const screenCode = normalizeIdInput(incomingRecord.screenCode);
            const existingIndex = merged.findIndex(
              (record) =>
                (record.backendSiteId &&
                  record.backendSiteId === incomingRecord.backendSiteId) ||
                (!record.backendSiteId &&
                  normalizeIdInput(record.screenCode) === screenCode),
            );

            if (existingIndex >= 0) {
              merged[existingIndex] = mergeStage3RecordBySiteId(
                merged[existingIndex],
                incomingRecord,
              );
              return;
            }

            merged.push(incomingRecord);
          });

          return merged;
        });

        // A site persisted as Stage 3 must not remain in the live Stage 2 queue.
        const stage3ScreenCodes = new Set(
          backendStage3Records
            .map((record) => normalizeIdInput(record.screenCode))
            .filter(Boolean),
        );

        if (stage3ScreenCodes.size > 0) {
          setStage2Records((currentRecords) =>
            currentRecords.filter(
              (record) => !stage3ScreenCodes.has(normalizeIdInput(record.screenCode)),
            ),
          );
        }

      } catch (error) {
        if (!cancelled) {
          console.error("Unable to restore Stage 3 records from backend:", error);
        }
      }
    }

    restoreStage3Records();

    return () => {
      cancelled = true;
    };
  }, []);

  const addStage2Records = useCallback((records, persistedSites = []) => {
    const correctionScreenCode = normalizeIdInput(stage1CorrectionScreenCode);

    // Stage 1 Readiness is common for new and returned/existing sites.
    // A pending returned site must not block unrelated new Stage 1 records
    // from moving to Stage 2.
    const persistedBySiteId = new Map(
      (Array.isArray(persistedSites) ? persistedSites : [])
        .filter((site) => site?.id)
        .map((site) => [site.id, site]),
    );
    const acceptedRecords = (Array.isArray(records) ? records : []).map(
      (record) => {
        const persistedSite = persistedBySiteId.get(record?.backendSiteId);
        if (!persistedSite) return record;

        const hydratedRecord = mergeCorrectedStage1FieldsIntoStage2Record(
          buildStage2RecordFromBackendSite(persistedSite, persistedSites),
          record,
        );

        return {
          ...hydratedRecord,
          backendSiteId: persistedSite.id,
          processingStatus: persistedSite.processingStatus || "Stage 2",
          readinessStatus: persistedSite.readinessStatus || "Moved to Stage 2",
          stage1CorrectionReturn: record.stage1CorrectionReturn === true,
        };
      },
    );

    setStage2Records((currentRecords) => {
      const nextRecords = [...currentRecords];
      const seenIncomingSiteIds = new Set();

      acceptedRecords.forEach((record) => {
        const normalizedSiteId = normalizeIdInput(record.screenCode);
        const canonicalIdentity = record.backendSiteId || normalizedSiteId;

        if (!normalizedSiteId || seenIncomingSiteIds.has(canonicalIdentity)) {
          return;
        }

        seenIncomingSiteIds.add(canonicalIdentity);

        const existingIndex = nextRecords.findIndex(
          (currentRecord) =>
            (record.backendSiteId &&
              currentRecord.backendSiteId === record.backendSiteId) ||
            (!record.backendSiteId &&
              normalizeIdInput(currentRecord.screenCode) === normalizedSiteId),
        );

        if (existingIndex >= 0) {
          if (!record.stage1CorrectionReturn) {
            return;
          }

          nextRecords[existingIndex] = mergeCorrectedStage1FieldsIntoStage2Record(
            nextRecords[existingIndex],
            record,
          );
          return;
        }

        // A site that was returned from Stage 2 to Stage 1 must re-enter
        // Stage 2 with its own preserved installation snapshot. Use the
        // Screen Code history map first so Stage 1 edits cannot accidentally
        // turn it into a fresh Installation Pending record.
        const preservedReentrySnapshot =
          stage2ReentrySnapshotsRef.current.get(normalizedSiteId) ||
          (record.stage1CorrectionReturn &&
          correctionScreenCode &&
          normalizedSiteId === correctionScreenCode &&
          stage1CorrectionSnapshot &&
          normalizeIdInput(stage1CorrectionSnapshot?.screenCode) === normalizedSiteId
            ? stage1CorrectionSnapshot
            : null);

        if (preservedReentrySnapshot) {
          const restoredRecord = mergeCorrectedStage1FieldsIntoStage2Record(
            preservedReentrySnapshot,
            record,
          );

          if (preservedReentrySnapshot.stage3ReturnForCorrection === true) {
            restoredRecord.expenseCycleStartedAt = new Date().toISOString();
            delete restoredRecord.stage3ReturnForCorrection;
          }

          nextRecords.push(restoredRecord);
          return;
        }

        nextRecords.push(record);
      });

      return nextRecords;
    });

    const correctionReturnedToStage2 =
      Boolean(correctionScreenCode) &&
      acceptedRecords.some(
        (record) =>
          record?.stage1CorrectionReturn === true &&
          normalizeIdInput(record?.screenCode) === correctionScreenCode,
      );

    acceptedRecords.forEach((record) => {
      const normalizedSiteId = normalizeIdInput(record?.screenCode);
      if (normalizedSiteId && stage2ReentrySnapshotsRef.current.has(normalizedSiteId)) {
        stage2ReentrySnapshotsRef.current.delete(normalizedSiteId);
      }
    });

    if (correctionReturnedToStage2) {
      setStage1CorrectionScreenCode("");
      setStage1CorrectionSnapshot(null);
      setStage1CorrectionRecord(null);
    }

    setStage2FocusRecordId("");
    setActivePage("installations");
  }, [stage1CorrectionScreenCode, stage1CorrectionSnapshot]);

  const handleReturnStage2RecordToStage1 = useCallback(async (record) => {
    if (!record) {
      return;
    }

    const normalizedScreenCode = normalizeIdInput(record.screenCode);

    if (!normalizedScreenCode) {
      return;
    }

    // Preserve the complete Stage 2 record before removing it from the active
    // Stage 2 list. It will be reused if this Screen Code is later moved back
    // from Stage 1 to Stage 2.
    stage2ReentrySnapshotsRef.current.set(normalizedScreenCode, { ...record });

    const returnedAt = new Date().toISOString();
    const returnedBillingCode = normalizeIdInput(record.billingCode);
    const returnedComplexCode = normalizeIdInput(record.complexCode, {
      allowStandaloneBlank: true,
    });

    const buildCommercialReferenceFromRecord = (source) => {
      if (!source) return {};

      return {
        otfApplicable: source.otfApplicable,
        otfType: source.otfType,
        otfTaxMode: source.otfTaxMode,
        otfAmount: source.otfAmount,
        subscriptionType: source.subscriptionType,
        subscriptionMode: source.subscriptionMode,
        subscriptionFee: source.subscriptionFee,
        pricingMethod: source.pricingMethod || "Common",
        pricingGroups: Array.isArray(source.pricingGroups)
          ? source.pricingGroups.map((group) => ({ ...group }))
          : [],
        allocationRows: Array.isArray(source.allocationRows)
          ? source.allocationRows.map((row) => ({ ...row }))
          : [],
        incentiveApplicable: source.incentiveApplicable,
        salesEmployeeName: source.salesEmployeeName,
        companyId: source.companyId,
        commercialReferenceScreenCode: source.screenCode || "",
      };
    };

    // Use an already-loaded non-FoC sibling first so the UI return is instant.
    const inSessionCommercialSource =
      record?.foc === true
        ? [...stage2Records, ...stage1ReadinessRecords].find((candidate) => {
            const candidateScreenCode = normalizeIdInput(candidate?.screenCode);
            const candidateBillingCode = normalizeIdInput(candidate?.billingCode);
            const candidateComplexCode = normalizeIdInput(candidate?.complexCode, {
              allowStandaloneBlank: true,
            });
            const candidatePricingMethod = String(candidate?.pricingMethod || "")
              .trim()
              .toLowerCase();

            return (
              candidateScreenCode &&
              candidateScreenCode !== normalizedScreenCode &&
              candidateBillingCode === returnedBillingCode &&
              candidateComplexCode === returnedComplexCode &&
              candidate?.foc !== true &&
              candidatePricingMethod !== "foc"
            );
          }) || null
        : null;

    let commercialReference = buildCommercialReferenceFromRecord(
      inSessionCommercialSource,
    );

    const makeCorrectionSnapshot = (reference = {}) => ({
      ...record,
      ...reference,
      // Preserve the returned screen's identity and FoC state.
      foc: record?.foc === true,
      selected: false,
      processingStatus: "Stage 1",
      readinessStatus: "Ready",
      stage1CorrectionReturn: true,
      updatedAt: returnedAt,
      stage1ReturnedAt: returnedAt,
    });

    // Return to Stage 1 immediately. Backend lookup/enrichment happens after
    // the visible state change and must never block the user's navigation.
    let correctionSnapshot = makeCorrectionSnapshot(commercialReference);

    setStage1ReadinessRecords((currentRecords) =>
      upsertStage1Records(currentRecords, [correctionSnapshot]),
    );

    setStage1CorrectionRecord(correctionSnapshot);
    setStage1CorrectionScreenCode(normalizedScreenCode);
    setStage1CorrectionSnapshot(correctionSnapshot);

    setStage2Records((currentRecords) =>
      currentRecords.filter(
        (currentRecord) =>
          normalizeIdInput(currentRecord.screenCode) !== normalizedScreenCode,
      ),
    );

    setActivePage("site-entry");

    // For FoC returns, enrich the returned Stage 1 record from the current
    // non-FoC sibling commercial term in the backend. This also restores 1B
    // after F5, even if the sibling commercial data was not in Stage 2 memory.
    if (record?.foc === true && returnedBillingCode) {
      try {
        const backendSites = await fetchBackendSites();
        const siblingSite = backendSites.find((site) => {
          const siblingBillingCode = normalizeIdInput(site?.billingId);
          const siblingComplexCode = normalizeIdInput(site?.complexId, {
            allowStandaloneBlank: true,
          });
          const siblingScreenCode = normalizeIdInput(site?.siteId);
          const currentCommercial =
            Array.isArray(site?.commercialTerms) && site.commercialTerms.length > 0
              ? site.commercialTerms[0]
              : null;

          return (
            siblingScreenCode &&
            siblingScreenCode !== normalizedScreenCode &&
            siblingBillingCode === returnedBillingCode &&
            siblingComplexCode === returnedComplexCode &&
            site?.foc !== true &&
            Boolean(currentCommercial)
          );
        });

        if (siblingSite) {
          const siblingStage1 =
            siblingSite?.stage1Data && typeof siblingSite.stage1Data === "object"
              ? siblingSite.stage1Data
              : {};
          const currentCommercial = siblingSite.commercialTerms[0];
          const pricingRows = Array.isArray(currentCommercial?.pricingRows)
            ? currentCommercial.pricingRows
            : [];
          const savedPricingMethod = String(
            siblingStage1?.pricingMethod || "",
          ).trim();
          const usableSavedPricingMethod =
            savedPricingMethod && savedPricingMethod.toLowerCase() !== "foc"
              ? savedPricingMethod
              : "";

          commercialReference = {
            otfApplicable:
              siblingStage1?.otfApplicable ||
              currentCommercial?.otfApplicable ||
              "No",
            otfType:
              siblingStage1?.otfType ||
              mapOtfTypeFromBackend(currentCommercial?.otfType),
            otfTaxMode:
              siblingStage1?.otfTaxMode ||
              mapTaxModeFromBackend(currentCommercial?.otfTaxMode),
            otfAmount:
              siblingStage1?.otfAmount ??
              decimalInputValue(currentCommercial?.otfAmount),
            subscriptionType:
              siblingStage1?.subscriptionType ||
              mapSubscriptionTypeFromBackend(currentCommercial?.subscriptionType) ||
              "Fixed",
            subscriptionMode:
              siblingStage1?.subscriptionMode ||
              mapSubscriptionModeFromBackend(currentCommercial?.subscriptionMode) ||
              "Monthly",
            subscriptionFee:
              siblingStage1?.subscriptionFee ??
              decimalInputValue(currentCommercial?.subscriptionFee),
            pricingMethod:
              usableSavedPricingMethod ||
              mapPricingMethodFromBackend(currentCommercial?.pricingMode) ||
              "Common",
            pricingGroups: Array.isArray(siblingStage1?.pricingGroups)
              ? siblingStage1.pricingGroups.map((group) => ({ ...group }))
              : [],
            allocationRows: Array.isArray(siblingStage1?.allocationRows)
              ? siblingStage1.allocationRows.map((row) => ({ ...row }))
              : pricingRows.map((row) => ({
                  screenCode: row.screenCode || "",
                  pricingGroup: row.pricingGroup || "",
                  deviceCount: row.deviceCount ?? "",
                  otfAmount: decimalInputValue(row.otfAmount),
                  otfTaxMode: mapTaxModeFromBackend(row.otfTaxMode),
                  subscriptionMode: row.planMode || "",
                  subscriptionFee: decimalInputValue(row.planFee),
                })),
            incentiveApplicable:
              siblingStage1?.incentiveApplicable ??
              siblingSite?.incentiveApplicable ??
              false,
            salesEmployeeName:
              siblingStage1?.salesEmployeeName ||
              siblingSite?.incentiveEmployeeName ||
              "",
            companyId:
              siblingStage1?.companyId ||
              siblingSite?.incentiveEmployeeNumber ||
              "",
            commercialReferenceScreenCode: siblingSite.siteId || "",
          };

          correctionSnapshot = makeCorrectionSnapshot(commercialReference);

          setStage1ReadinessRecords((currentRecords) =>
            currentRecords.map((currentRecord) =>
              normalizeIdInput(currentRecord.screenCode) === normalizedScreenCode
                ? correctionSnapshot
                : currentRecord,
            ),
          );
          setStage1CorrectionSnapshot(correctionSnapshot);
        }
      } catch (error) {
        console.error("Unable to restore Stage 1 commercial reference", error);
      }
    }

    // Persist the workflow return so F5 cannot restore the site into Stage 2.
    const backendSiteId = record.backendSiteId;
    if (backendSiteId) {
      fetch(`${BILLING_API_BASE}/sites/${backendSiteId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          processingStatus: "Stage 1",
          readinessStatus: "Ready",
          // Keep the current Stage 1 projection authoritative while leaving
          // the historical Stage 2/Stage 3 snapshots untouched.
          stage1Data: {
            ...record,
            processingStatus: "Stage 1",
            readinessStatus: "Ready",
            selected: false,
          },
        }),
      }).catch((error) => {
        console.error("Unable to persist Stage 1 return:", error);
      });
    }
  }, [stage1ReadinessRecords, stage2Records]);

  const handleStage1CorrectionRecordConsumed = useCallback(() => {
    // Clear only the UI hand-off object after SiteEntry has loaded it.
    // Keep stage1CorrectionScreenCode + stage1CorrectionSnapshot until the
    // corrected screen successfully re-enters Stage 2.
    setStage1CorrectionRecord(null);
  }, []);

  const moveStage2RecordsToStage3 = useCallback(
    (records) => {
      const incomingSiteIds = new Set();
      const duplicateSiteId = records.find((record) => {
        const screenCode = normalizeIdInput(record.screenCode);
        if (!screenCode) {
          return true;
        }

        if (incomingSiteIds.has(screenCode)) {
          return true;
        }

        incomingSiteIds.add(screenCode);
        return false;
      });

      if (duplicateSiteId) {
        return false;
      }

      const existingStage3SiteIds = new Set(
        stage3Records
          .map((record) => normalizeIdInput(record?.screenCode))
          .filter(Boolean),
      );

      const alreadyActiveInStage3 = records.find((record) =>
        existingStage3SiteIds.has(normalizeIdInput(record?.screenCode)),
      );

      if (alreadyActiveInStage3) {
        alert(
          `${alreadyActiveInStage3.screenCode} already exists in Stage 3. ` +
            "A Screen Code can exist in only one active workflow stage at a time.",
        );
        return false;
      }

      setStage3Records((currentRecords) => [
        ...currentRecords.filter(
          (currentRecord) =>
            !records.some(
              (record) =>
                normalizeIdInput(record.screenCode) ===
                normalizeIdInput(currentRecord.screenCode),
            ),
        ),
        ...records.map((record) => {
          const incomingStage3Record = createStage3Record(record);
          const existingRecord = currentRecords.find(
            (currentRecord) =>
              normalizeIdInput(currentRecord.screenCode) ===
              normalizeIdInput(record.screenCode),
          );

          return existingRecord
            ? mergeStage3RecordBySiteId(existingRecord, incomingStage3Record)
            : incomingStage3Record;
        }),
      ]);

      // Persist the workflow handoff. Without this, a browser refresh reloads
      // the same site as Stage 2 and it reappears in Billing Readiness.
      records.forEach((record) => {
        const backendSiteId = record?.backendSiteId;
        if (!backendSiteId) {
          console.warn(
            `Unable to persist Stage 3 handoff for ${record?.screenCode || "site"}: backend site id is missing.`,
          );
          return;
        }

        fetch(`${BILLING_API_BASE}/sites/${backendSiteId}`, {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            processingStatus: "Stage 3",
            readinessStatus: "Moved to Stage 3",
            stage3Data: createStage3Record(record),
          }),
        })
          .then(async (response) => {
            if (response.ok) return;
            const body = await response.json().catch(() => ({}));
            throw new Error(
              body?.message ||
                body?.error ||
                `Unable to persist Stage 3 handoff (${response.status}).`,
            );
          })
          .catch((error) => {
            console.error(
              `Unable to persist Stage 3 handoff for ${record?.screenCode || "site"}:`,
              error,
            );
          });
      });

      setActivePage("billings");
      return true;
    },
    [stage3Records],
  );

  const returnStage3RecordsToStage2 = useCallback(
    async (records) => {
      const incomingSiteIds = new Set();
      const duplicateWithinReturn = records.find((record) => {
        const screenCode = normalizeIdInput(record.screenCode);
        if (!screenCode) {
          return true;
        }

        if (incomingSiteIds.has(screenCode)) {
          return true;
        }

        incomingSiteIds.add(screenCode);
        return false;
      });

      if (duplicateWithinReturn) {
        return false;
      }

      const recordsWithoutBackendIds = records.filter(
        (record) => !record?.backendSiteId,
      );

      if (recordsWithoutBackendIds.length > 0) {
        console.error(
          "Unable to return Stage 3 records to Stage 2: backend Site ID is missing.",
        );
        return false;
      }

      try {
        await Promise.all(
          records.map(async (record) => {
            const response = await fetch(
              `${BILLING_API_BASE}/sites/${record.backendSiteId}`,
              {
                method: "PATCH",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                  processingStatus: "Stage 2",
                  readinessStatus: "Moved to Stage 2",
                  stage2Data: record.stage2Data || undefined,
                }),
              },
            );

            if (!response.ok) {
              const body = await response.json().catch(() => ({}));
              throw new Error(
                body?.message ||
                  body?.error ||
                  `Unable to persist Stage 2 return (${response.status}).`,
              );
            }
          }),
        );
      } catch (error) {
        console.error("Unable to persist Stage 2 return:", error);
        return false;
      }

      const recordSiteIds = new Set(
        records.map((record) => normalizeIdInput(record.screenCode)).filter(Boolean),
      );
      const expenseCycleStartedAt = new Date().toISOString();

      setStage2Records((currentRecords) => {
        const nextRecords = currentRecords.filter(
          (record) => !recordSiteIds.has(normalizeIdInput(record.screenCode)),
        );

        records.forEach((record) => {
          const normalizedSiteId = normalizeIdInput(record.screenCode);
          const existingIndex = nextRecords.findIndex(
            (currentRecord) =>
              normalizeIdInput(currentRecord.screenCode) === normalizedSiteId,
          );

          const restoredRecord = {
            ...record,
            selected: false,
            readinessStatus: record.readinessStatus || "Ready",
            stage3BillingStartDateSnapshot: normalizeDateValue(
              record.billingStartDate,
            ),
            billingVerificationStatus: "",
            billingVerifiedAt: "",
            closureStatus: "Open",
            closureSelected: false,
            expenseCycleStartedAt,
          };

          if (existingIndex >= 0) {
            nextRecords[existingIndex] = mergeStage2RecordBySiteId(
              nextRecords[existingIndex],
              restoredRecord,
            );
            return;
          }

          nextRecords.push(restoredRecord);
        });

        return nextRecords;
      });

      setStage3Records((currentRecords) =>
        currentRecords.filter(
          (record) => !recordSiteIds.has(normalizeIdInput(record.screenCode)),
        ),
      );

      setStage2FocusRecordId(records[0]?.recordId || "");
      setActivePage("installations");
      return true;
    },
    [],
  );

  const routeMouStatusRecordToStage1 = useCallback((record, workflowType) => {
    if (!record) {
      return false;
    }

    const normalizedScreenCode = normalizeIdInput(record.screenCode);

    if (!normalizedScreenCode) {
      return false;
    }

    const normalizedWorkflowType =
      String(workflowType || "").trim() || "Addendum";

    const billingCode = normalizeIdInput(record.billingCode);
    const complexCode = normalizeIdInput(record.complexCode, {
      allowStandaloneBlank: true,
    });

    const isComplexRenewal =
      normalizedWorkflowType === "Renewal" && Boolean(complexCode);

    // Renewal scope:
    // - Standalone => selected site only.
    // - Complex => every active screen under the same Billing Code + Complex Code.
    // Addendum remains selective and continues through the existing Addendum flow.
    const sourceRecords = [...stage2Records, ...stage3Records];
    const renewalGroupRecords = isComplexRenewal
      ? sourceRecords.filter((candidate) => {
          const candidateBillingCode = normalizeIdInput(candidate?.billingCode);
          const candidateComplexCode = normalizeIdInput(candidate?.complexCode, {
            allowStandaloneBlank: true,
          });

          return (
            candidateBillingCode === billingCode &&
            candidateComplexCode === complexCode
          );
        })
      : [record];

    const uniqueRenewalRecords = Array.from(
      renewalGroupRecords.reduce((map, candidate) => {
        const siteId = normalizeIdInput(candidate?.screenCode);

        if (siteId && !map.has(siteId)) {
          map.set(siteId, candidate);
        }

        return map;
      }, new Map()).values(),
    );

    if (isComplexRenewal && uniqueRenewalRecords.length === 0) {
      return false;
    }

    if (isComplexRenewal) {
      const confirmed = window.confirm(
        `Complex Renewal will update all ${uniqueRenewalRecords.length} screen${
          uniqueRenewalRecords.length === 1 ? "" : "s"
        } under Complex ${complexCode}. Continue?`,
      );

      if (!confirmed) {
        return false;
      }
    }

    const renewalGroupId = isComplexRenewal
      ? `renewal-${billingCode}-${complexCode}`
      : record.stage1GroupId || "";

    const renewalGroupRows = isComplexRenewal
      ? uniqueRenewalRecords.map((candidate) => ({
          billingCode: normalizeIdInput(candidate.billingCode),
          complexCode: normalizeIdInput(candidate.complexCode, {
            allowStandaloneBlank: true,
          }),
          billingName: candidate.billingName || "",
          screenCode: normalizeIdInput(candidate.screenCode),
          screenName: candidate.screenName || "",
          location: candidate.location || "",
          state: candidate.state || "",
          siteType: candidate.siteType || "Complex",
          selected: false,
          status: candidate.status || "Complete",
        }))
      : Array.isArray(record.stage1GroupRows)
        ? record.stage1GroupRows.map((row) => ({ ...row, selected: false }))
        : [];

    const buildAgreementSnapshot = (candidate) => ({
      ...candidate,
      stage1GroupId:
        isComplexRenewal ? renewalGroupId : candidate.stage1GroupId,
      stage1GroupRows:
        isComplexRenewal ? renewalGroupRows : candidate.stage1GroupRows,
      stage1CorrectionReturn: true,
      complexRenewal: isComplexRenewal,
      renewalScope: isComplexRenewal ? "Complex" : "Standalone",
      correctionType: normalizedWorkflowType,
      correctionRoute:
        normalizedWorkflowType === "Addendum"
          ? "MoU Status → Stage 1 Addendum"
          : isComplexRenewal
            ? "MoU Status → Stage 1 Complex Renewal"
            : "MoU Status → Stage 1 Renewal",
      agreementType:
        normalizedWorkflowType === "Addendum"
          ? "Addendum"
          : candidate.agreementType || "MoU",
      previousAgreementSnapshot:
        candidate.previousAgreementSnapshot || {
          agreementType: candidate.agreementType || "MoU",
          mouStartDate: candidate.mouStartDate || "",
          mouEndDate: candidate.mouEndDate || "",
          subscriptionType: candidate.subscriptionType || "",
          subscriptionMode: candidate.subscriptionMode || "",
          subscriptionFee: candidate.subscriptionFee || "",
          otfApplicable: candidate.otfApplicable || "No",
          otfType: candidate.otfType || "",
          otfAmount: candidate.otfAmount || "",
          pricingMethod: candidate.pricingMethod || "Common",
          pricingGroups: Array.isArray(candidate.pricingGroups)
            ? candidate.pricingGroups.map((group) => ({ ...group }))
            : [],
          salesEmployeeName: candidate.salesEmployeeName || "",
          companyId: candidate.companyId || "",
        },
    });

    const agreementSnapshots = uniqueRenewalRecords.map(buildAgreementSnapshot);
    const agreementSnapshot =
      agreementSnapshots.find(
        (candidate) =>
          normalizeIdInput(candidate.screenCode) === normalizedScreenCode,
      ) ||
      agreementSnapshots[0] ||
      buildAgreementSnapshot(record);

    // Embed the full complex renewal group in the hand-off object. SiteEntry
    // expands these into one Stage 1 Readiness row per screen.
    const stage1HandoffRecord = isComplexRenewal
      ? {
          ...agreementSnapshot,
          renewalGroupRecords: agreementSnapshots,
        }
      : agreementSnapshot;

    setStage1CorrectionRecord(stage1HandoffRecord);

    // A complex renewal is intentionally not a single-screen correction.
    // Leave the correction screen blank so all complex screens can be edited/saved together.
    setStage1CorrectionScreenCode(
      isComplexRenewal ? "" : normalizedScreenCode,
    );
    setStage1CorrectionSnapshot(
      isComplexRenewal ? null : agreementSnapshot,
    );

    const renewalSiteIds = new Set(
      agreementSnapshots
        .map((candidate) => normalizeIdInput(candidate.screenCode))
        .filter(Boolean),
    );

    setStage2Records((currentRecords) =>
      currentRecords.filter(
        (currentRecord) =>
          !renewalSiteIds.has(normalizeIdInput(currentRecord.screenCode)),
      ),
    );

    setStage3Records((currentRecords) =>
      currentRecords.filter(
        (currentRecord) =>
          !renewalSiteIds.has(normalizeIdInput(currentRecord.screenCode)),
      ),
    );

    setStage2FocusRecordId("");
    setActivePage("site-entry");
    return true;
  }, [stage2Records, stage3Records]);

  const routeStage3RecordToStage1Correction = useCallback(async (record) => {
    if (!record) {
      return false;
    }

    const normalizedScreenCode = normalizeIdInput(record.screenCode);

    if (!normalizedScreenCode) {
      return false;
    }

    // Preserve the previous Stage 2 installation state before routing this
    // Stage 3 record back to Stage 1. If the same screen is later moved from
    // Stage 1 to Stage 2, addStage2Records restores this snapshot instead of
    // treating the site as a fresh Installation Pending record.
    stage2ReentrySnapshotsRef.current.set(normalizedScreenCode, {
      ...record,
      stage3ReturnForCorrection: true,
    });

    const returnedAt = new Date().toISOString();
    const correctionSnapshot = {
      ...record,
      selected: false,
      processingStatus: "Stage 1",
      readinessStatus: "Ready",
      stage1CorrectionReturn: true,
      correctionRequestedAt: record.correctionRequestedAt || returnedAt,
      stage1ReturnedAt: returnedAt,
      updatedAt: returnedAt,
    };

    // Persist the Stage 3 -> Stage 1 return so F5 cannot restore this screen
    // into Stage 3. Existing Stage 1 commercial snapshots remain untouched.
    const correctionEventId = record.correctionEventId || crypto.randomUUID();
    let backendSiteId = record?.backendSiteId;
    if (!backendSiteId) {
      try {
        const backendSites = await fetchBackendSites();
        backendSiteId = backendSites.find(
          (site) => normalizeIdInput(site?.siteId) === normalizedScreenCode,
        )?.id;
      } catch (error) {
        console.error("Unable to resolve the canonical Site ID for correction:", error);
      }
    }
    if (backendSiteId) {
      try {
        const correctionResponse = await fetch(
          `${BILLING_API_BASE}/sites/${backendSiteId}/corrections`,
          {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              ...(record.correctionAuditId
                ? { auditId: record.correctionAuditId }
                : {}),
              workflowType: "RECORD_CORRECTION",
              eventId: correctionEventId,
              correctionFields: [
                ...(Array.isArray(record.correctionFields)
                  ? record.correctionFields
                  : []),
                "processingStatus",
                "readinessStatus",
              ],
              reason: record.correctionReason || "Returned to Stage 1",
              previousValues: {
                processingStatus: "Stage 3",
                readinessStatus: "Moved to Stage 3",
              },
              newValues: {
                processingStatus: "Stage 1",
                readinessStatus: "Ready",
              },
              completeReturn: true,
              stage1Data: {
                ...record,
                correctionEventId,
                processingStatus: "Stage 1",
                readinessStatus: "Ready",
                selected: false,
              },
            }),
          },
        );
        if (!correctionResponse.ok) {
          throw new Error(`Unable to record Stage 1 return (${correctionResponse.status}).`);
        }
      } catch (error) {
        console.error(
          `Unable to persist Stage 1 return for ${record?.screenCode || "site"}:`,
          error,
        );
        return false;
      }
    } else {
      console.warn(
        `Unable to persist Stage 1 return for ${record?.screenCode || "site"}: backend site id is missing.`,
      );
      return false;
    }

    // The UI projection changes only after correction audit and Site update
    // have committed successfully together.
    const persistedCorrectionSnapshot = {
      ...correctionSnapshot,
      correctionEventId,
    };
    setStage1ReadinessRecords((currentRecords) =>
      upsertStage1Records(currentRecords, [persistedCorrectionSnapshot]),
    );
    setStage1CorrectionRecord(persistedCorrectionSnapshot);
    setStage1CorrectionScreenCode(normalizedScreenCode);
    setStage1CorrectionSnapshot(persistedCorrectionSnapshot);
    setStage2Records((currentRecords) =>
      currentRecords.filter(
        (currentRecord) =>
          normalizeIdInput(currentRecord.screenCode) !== normalizedScreenCode,
      ),
    );
    setStage3Records((currentRecords) =>
      currentRecords.filter(
        (currentRecord) =>
          normalizeIdInput(currentRecord.screenCode) !== normalizedScreenCode,
      ),
    );

    setStage2FocusRecordId("");
    setActivePage("site-entry");
    return true;
  }, []);

  const mouStatusRecords = [
    ...stage2Records,
    ...stage3Records,
  ].reduce((rows, record) => {
    const screenCode = normalizeIdInput(record?.screenCode);

    if (!screenCode) {
      return rows;
    }

    const existingIndex = rows.findIndex(
      (row) => normalizeIdInput(row?.screenCode) === screenCode,
    );

    if (existingIndex >= 0) {
      rows[existingIndex] = record;
      return rows;
    }

    rows.push(record);
    return rows;
  }, []);

  const handleCreateAddendumFromMouStatus = useCallback(
    (record) => routeMouStatusRecordToStage1(record, "Addendum"),
    [routeMouStatusRecordToStage1],
  );

  const handleOpenRenewalFromMouStatus = useCallback(
    (record) => routeMouStatusRecordToStage1(record, "Renewal"),
    [routeMouStatusRecordToStage1],
  );

  function renderActivePage() {
    if (activePage === "site-entry") {
      return (
        <SiteEntry
          readinessRecords={stage1ReadinessRecords}
          setReadinessRecords={setStage1ReadinessRecords}
          existingStage2Records={stage2Records}
          existingStage3Records={currentStage3Records}
          setStage2Records={setStage2Records}
          setStage3Records={setStage3Records}
  onMoveToStage2={addStage2Records}
  returnedStage1Record={stage1CorrectionRecord}
  onReturnedStage1RecordConsumed={handleStage1CorrectionRecordConsumed}
/>
      );
    }

    if (activePage === "installations") {
      return (
        <Installations
          stage2Records={stage2Records}
          setStage2Records={setStage2Records}
          referenceStage3Records={currentStage3Records}
          onMoveToStage3={moveStage2RecordsToStage3}
          onReturnToStage1={handleReturnStage2RecordToStage1}
          focusRecordId={stage2FocusRecordId}
          onFocusRecordConsumed={() => setStage2FocusRecordId("")}
        />
      );
    }

    if (activePage === "billings") {
      return (
        <Billings
          stage1Records={stage1ReadinessRecords}
          stage2Records={stage2Records}
          setStage2Records={setStage2Records}
          stage3Records={currentStage3Records}
          canonicalSiteRecords={canonicalSiteRecords}
          setStage3Records={setStage3Records}
          billingRecords={billingRecords}
          setBillingRecords={setBillingRecords}
          commonComplexAllocations={commonComplexAllocations}
          setCommonComplexAllocations={setCommonComplexAllocations}
          onReturnToStage2={returnStage3RecordsToStage2}
          onReturnCommercialCorrectionToStage1={routeStage3RecordToStage1Correction}
        />
      );
    }

    if (activePage === "incentive") {
      return (
        <Incentive
          stage3Records={currentStage3Records}
          billingRecords={billingRecords}
          incentiveStates={incentiveStates}
          setIncentiveStates={setIncentiveStates}
        />
      );
    }

    if (activePage === "mou-status") {
      return (
        <MouStatus
          records={mouStatusRecords}
          onCreateAddendum={handleCreateAddendumFromMouStatus}
          onOpenRenewal={handleOpenRenewalFromMouStatus}
        />
      );
    }

    if (activePage === "history") {
      return (
        <HistoryTracking
          stage1Records={stage1ReadinessRecords}
          stage2Records={stage2Records}
          stage3Records={currentStage3Records}
          billingRecords={billingRecords}
          incentiveStates={incentiveStates}
        />
      );
    }

    if (activePage === "audit") {
      return (
        <AuditHistory
          stage1Records={stage1ReadinessRecords}
          stage2Records={stage2Records}
          stage3Records={currentStage3Records}
          billingRecords={billingRecords}
          incentiveStates={incentiveStates}
        />
      );
    }

    if (activePage === "downloads") {
      return (
        <Downloads
          stage1Records={stage1ReadinessRecords}
          stage2Records={stage2Records}
          stage3Records={currentStage3Records}
          billingRecords={billingRecords}
          incentiveStates={incentiveStates}
        />
      );
    }

    if (activePage === "settings") {
      return (
        <Settings
          erpMappings={erpMappings}
          setErpMappings={setErpMappings}
        />
      );
    }

    return (
      <div className="page-placeholder">
        <p className="page-placeholder__label">Current Module</p>
        <h3>{activePage.replace("-", " ")}</h3>
        <p>This module will be built after Site Entry.</p>
      </div>
    );
  }

  return (
    <AppLayout activePage={activePage} onPageChange={setActivePage}>
      {renderActivePage()}
    </AppLayout>
  );
}

export default App;
