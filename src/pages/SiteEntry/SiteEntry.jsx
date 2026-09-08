import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import SiteCreation from "./components/SiteCreation";
import MouSubscription from "./components/MouSubscription";
import "./SiteEntry.css";
import {
  getDisplayComplexCode,
  getIdValidationMessage,
  isStandaloneComplexCode,
  isValidIdInput,
  normalizeIdInput,
} from "../../utils/idValidation";
import {
  getCurrentWorkflowStage,
  isStage1ReadyForSelection,
} from "../../utils/workflowStage";

const GUIDED_SCROLL_DURATION_MS = 800;
const GUIDED_SCROLL_TOP_OFFSET_PX = 24;
const BILLING_API_BASE =
  import.meta.env.VITE_BILLING_API_BASE || "http://localhost:4100/api";

function toNumberOrNull(value) {
  if (value === "" || value === null || value === undefined) {
    return null;
  }

  const numberValue = Number(String(value).replace(/,/g, ""));
  return Number.isFinite(numberValue) ? numberValue : null;
}

function mapPricingModeForApi(value) {
  const normalized = normalizeValue(value).toLowerCase();

  if (normalized === "site-wise" || normalized === "sitewise") return "sitewise";
  if (normalized === "allocation model" || normalized === "allocation") return "allocation";
  return "common";
}

function mapSubscriptionTypeForApi(value) {
  if (!value || value === "Not Applicable") return null;
  return value === "Variable" ? "Vary" : value;
}

function mapOtfTypeFromApi(value) {
  if (value === "NonRefundable") return "Non-Refundable";
  return value || "";
}

function mapTaxModeFromApi(value) {
  if (value === "IncludingTax") return "Including Tax";
  if (value === "ExcludingTax") return "Excluding Tax";
  if (value === "NotApplicable") return "Not Applicable";
  return value || "";
}

function mapSubscriptionTypeFromApi(value) {
  if (value === "Vary") return "Variable";
  return value || "";
}

function mapSubscriptionModeFromApi(value) {
  if (value === "HalfYearly") return "Half-Yearly";
  return value || "";
}

function decimalToInputValue(value) {
  if (value === null || value === undefined || value === "") return "";
  return String(value);
}

function getReadinessRecordIdentity(record) {
  const backendSiteId = normalizeValue(record?.backendSiteId);
  const screenCode = normalizeIdInput(record?.screenCode);

  return {
    backendSiteId,
    screenCode,
  };
}

function upsertReadinessRecords(currentRecords, incomingRecords) {
  const incoming = Array.isArray(incomingRecords) ? incomingRecords : [];
  const incomingIdentities = incoming.map(getReadinessRecordIdentity);
  const retained = currentRecords.filter((record) => {
    const identity = getReadinessRecordIdentity(record);

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
    const identity = getReadinessRecordIdentity(record);
    const alreadyAdded = uniqueIncoming.some((existingRecord) => {
      const existingIdentity = getReadinessRecordIdentity(existingRecord);
      return (
        (identity.backendSiteId &&
          identity.backendSiteId === existingIdentity.backendSiteId) ||
        (identity.screenCode && identity.screenCode === existingIdentity.screenCode)
      );
    });

    if (!alreadyAdded) {
      uniqueIncoming.push(record);
    }
  });

  return [...retained, ...uniqueIncoming];
}

function mapSiteTypeForApi(value) {
  const normalized = normalizeValue(value).toLowerCase();
  return normalized === "complex" || normalized === "multiplex"
    ? "Complex"
    : "Single";
}

async function apiRequest(path, options = {}) {
  const response = await fetch(`${BILLING_API_BASE}${path}`, {
    headers: {
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
    ...options,
  });

  const body = await response.json().catch(() => ({}));

  if (!response.ok) {
    const fieldErrors = body?.details?.fieldErrors || {};
    const formErrors = Array.isArray(body?.details?.formErrors)
      ? body.details.formErrors
      : [];

    const validationMessages = [
      ...formErrors,
      ...Object.entries(fieldErrors).flatMap(([field, messages]) =>
        (Array.isArray(messages) ? messages : [messages])
          .filter(Boolean)
          .map((message) => `${field}: ${message}`),
      ),
    ];

    const baseMessage =
      body?.message ||
      body?.error ||
      `Backend request failed with status ${response.status}.`;

    throw new Error(
      validationMessages.length > 0
        ? `${baseMessage} - ${validationMessages.join("; ")}`
        : baseMessage,
    );
  }

  return body?.data;
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

function normalizeValue(value) {
  return String(value || "").trim();
}

function getStage1GroupId(record) {
  return record?.stage1GroupId || record?.savedAt || record?.recordId || "";
}

function getPrimaryCommercialTerm(site) {
  return Array.isArray(site?.commercialTerms) && site.commercialTerms.length > 0
    ? site.commercialTerms[0]
    : null;
}

function getLatestApprovedCommonPriceChange(site, sites = [], referenceDate = new Date()) {
  const billingCode = normalizeIdInput(site?.billingId);
  const complexCode = normalizeIdInput(site?.complexId, {
    allowStandaloneBlank: true,
  });
  const targetScreenCode = normalizeIdInput(site?.siteId);
  if (!billingCode || !complexCode || site?.foc === true) return null;

  const referenceDay = new Date(referenceDate);
  referenceDay.setHours(23, 59, 59, 999);
  return (Array.isArray(sites) ? sites : [])
    .filter(
      (candidate) =>
        candidate?.foc !== true &&
        normalizeIdInput(candidate?.billingId) === billingCode &&
        normalizeIdInput(candidate?.complexId, { allowStandaloneBlank: true }) ===
          complexCode,
    )
    .flatMap((candidate) => {
      const snapshot =
        candidate?.stage1Data && typeof candidate.stage1Data === "object"
          ? candidate.stage1Data
          : {};
      const inSelectedScope =
        normalizeIdInput(candidate?.siteId) === targetScreenCode ||
        (Array.isArray(snapshot.stage1GroupRows) &&
          snapshot.stage1GroupRows.some(
            (row) => normalizeIdInput(row?.screenCode) === targetScreenCode,
          ));
      if (!inSelectedScope) return [];

      const stage3 =
        candidate?.stage3Data && typeof candidate.stage3Data === "object"
          ? candidate.stage3Data
          : {};
      return (Array.isArray(stage3.priceChangeHistory)
        ? stage3.priceChangeHistory
        : []
      )
        .filter((entry) => {
          const effectiveDate = new Date(`${normalizeValue(entry?.effectiveDate)}T00:00:00`);
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
          _approvedAt:
            Date.parse(entry.validatedAt || entry.savedAndSubmittedAt || "") || 0,
        }));
    })
    .sort(
      (left, right) =>
        right._effectiveAt - left._effectiveAt ||
        right._approvedAt - left._approvedAt,
    )[0] || null;
}

function getLatestCommonCommercialTerm(site, sites = []) {
  const billingCode = normalizeIdInput(site?.billingId);
  const complexCode = normalizeIdInput(site?.complexId, {
    allowStandaloneBlank: true,
  });
  const targetScreenCode = normalizeIdInput(site?.siteId);

  if (!billingCode || !complexCode || site?.foc === true) {
    return getPrimaryCommercialTerm(site);
  }

  const candidates = (Array.isArray(sites) ? sites : [])
    .filter(
      (candidate) =>
        candidate?.foc !== true &&
        normalizeIdInput(candidate?.billingId) === billingCode &&
        normalizeIdInput(candidate?.complexId, { allowStandaloneBlank: true }) ===
          complexCode,
    )
    .flatMap((candidate) => {
      const candidateSnapshot =
        candidate?.stage1Data && typeof candidate.stage1Data === "object"
          ? candidate.stage1Data
          : {};
      const scopeRows = Array.isArray(candidateSnapshot.stage1GroupRows)
        ? candidateSnapshot.stage1GroupRows
        : [];
      const inSelectedScope =
        normalizeIdInput(candidate?.siteId) === targetScreenCode ||
        scopeRows.some(
          (row) => normalizeIdInput(row?.screenCode) === targetScreenCode,
        );

      if (!inSelectedScope) return [];

      return (Array.isArray(candidate?.commercialTerms)
        ? candidate.commercialTerms
        : []
      )
        .filter(
          (term) =>
            term?.isCurrent !== false && term?.pricingMode === "common",
        )
        .map((term) => ({
          ...term,
          _sourceUpdatedAt:
            Date.parse(term.updatedAt || term.createdAt || "") || 0,
        }));
    })
    .sort(
      (left, right) =>
        right._sourceUpdatedAt - left._sourceUpdatedAt ||
        (Number(right.version) || 0) - (Number(left.version) || 0),
    );

  const baseTerm = candidates[0] || getPrimaryCommercialTerm(site);
  const latestPriceChange = getLatestApprovedCommonPriceChange(site, sites);
  return latestPriceChange
    ? {
        ...baseTerm,
        subscriptionFee: latestPriceChange.newFee,
        subscriptionMode: latestPriceChange.newMode || baseTerm?.subscriptionMode,
        updatedAt: latestPriceChange.validatedAt || baseTerm?.updatedAt,
        _sourceUpdatedAt: latestPriceChange._approvedAt || latestPriceChange._effectiveAt,
        _source: "Price Change & Reallocation",
        _effectiveDate: latestPriceChange.effectiveDate,
      }
    : baseTerm;
}

function hasBillingActuallyStarted(site) {
  const snapshots = [site?.stage3Data, site?.stage2Data, site?.stage1Data];
  return snapshots.some((snapshot) => {
    if (!snapshot || typeof snapshot !== "object" || Array.isArray(snapshot)) {
      return false;
    }

    const verificationStatus = normalizeValue(
      snapshot.billingVerificationStatus || snapshot.submissionStatus,
    ).toLowerCase();
    return (
      ["submitted to billing team", "sent to billing team", "completed"].includes(
        verificationStatus,
      ) ||
      Boolean(snapshot.firstTimeBillingRecord?.billingRecordId) ||
      Boolean(snapshot.invoiceNumber) ||
      Boolean(snapshot.invoiceEntries?.length)
    );
  });
}

function buildCanonicalStage1EditRecord(record, backendSiteRecords = []) {
  if (!record) {
    return null;
  }

  const sites = Array.isArray(backendSiteRecords) ? backendSiteRecords : [];
  const recordBackendSiteId = normalizeValue(record?.backendSiteId);
  const recordScreenCode = normalizeIdInput(record?.screenCode);
  const backendSite =
    sites.find((site) => normalizeValue(site?.id) === recordBackendSiteId) ||
    sites.find((site) => normalizeIdInput(site?.siteId) === recordScreenCode) ||
    null;

  if (!backendSite) {
    return record;
  }

  const snapshot =
    backendSite?.stage1Data && typeof backendSite.stage1Data === "object"
      ? backendSite.stage1Data
      : {};
  const currentCommercial = getLatestCommonCommercialTerm(backendSite, sites);
  const useLatestCommonCommercial =
    backendSite.foc !== true && currentCommercial?.pricingMode === "common";
  const snapshotHas = (key) =>
    Object.prototype.hasOwnProperty.call(snapshot, key) &&
    snapshot[key] !== null &&
    snapshot[key] !== undefined &&
    snapshot[key] !== "";
  const pickValue = (...values) =>
    values.find((value) => value !== null && value !== undefined && value !== "") ??
    "";

  return {
    ...record,
    ...snapshot,
    backendSiteId: backendSite.id || record.backendSiteId || "",
    recordId: snapshot.recordId || record.recordId || backendSite.id,
    billingCode: pickValue(backendSite.billingId, snapshot.billingCode, record.billingCode),
    complexCode: pickValue(
      backendSite.complexId,
      snapshot.complexCode,
      record.complexCode,
    ),
    screenCode: pickValue(backendSite.siteId, snapshot.screenCode, record.screenCode),
    screenName: pickValue(
      backendSite.screenName,
      snapshot.screenName,
      record.screenName,
    ),
    location: pickValue(backendSite.location, snapshot.location, record.location),
    state: pickValue(backendSite.state, snapshot.state, record.state),
    foc: backendSite.foc === true ? true : Boolean(snapshot.foc ?? record.foc),
    siteType: pickValue(snapshot.siteType, record.siteType, backendSite.siteType),
    processingStatus: pickValue(
      backendSite.processingStatus,
      snapshot.processingStatus,
      record.processingStatus,
      "New",
    ),
    readinessStatus: pickValue(
      backendSite.readinessStatus,
      snapshot.readinessStatus,
      record.readinessStatus,
      "Ready",
    ),
    otfApplicable: snapshotHas("otfApplicable")
      ? snapshot.otfApplicable
      : pickValue(currentCommercial?.otfApplicable, record.otfApplicable, "No"),
    otfType: snapshotHas("otfType")
      ? snapshot.otfType
      : pickValue(mapOtfTypeFromApi(currentCommercial?.otfType), record.otfType),
    otfTaxMode: snapshotHas("otfTaxMode")
      ? snapshot.otfTaxMode
      : pickValue(
          mapTaxModeFromApi(currentCommercial?.otfTaxMode),
          record.otfTaxMode,
        ),
    otfAmount: snapshotHas("otfAmount")
      ? snapshot.otfAmount
      : decimalToInputValue(currentCommercial?.otfAmount) || record.otfAmount || "",
    subscriptionType: useLatestCommonCommercial
      ? mapSubscriptionTypeFromApi(currentCommercial?.subscriptionType)
      : snapshotHas("subscriptionType")
      ? snapshot.subscriptionType
      : pickValue(
          mapSubscriptionTypeFromApi(currentCommercial?.subscriptionType),
          record.subscriptionType,
          "Fixed",
        ),
    subscriptionMode: useLatestCommonCommercial
      ? mapSubscriptionModeFromApi(currentCommercial?.subscriptionMode)
      : snapshotHas("subscriptionMode")
      ? snapshot.subscriptionMode
      : pickValue(
          mapSubscriptionModeFromApi(currentCommercial?.subscriptionMode),
          record.subscriptionMode,
          "Monthly",
        ),
    subscriptionFee: useLatestCommonCommercial
      ? decimalToInputValue(currentCommercial?.subscriptionFee)
      : snapshotHas("subscriptionFee")
      ? snapshot.subscriptionFee
      : decimalToInputValue(currentCommercial?.subscriptionFee) || record.subscriptionFee || "",
    pricingMethod: snapshotHas("pricingMethod")
      ? snapshot.pricingMethod
      : pickValue(
          normalizeValue(currentCommercial?.pricingMode) === "sitewise"
            ? "Site-wise"
            : normalizeValue(currentCommercial?.pricingMode) === "allocation"
              ? "Allocation Model"
              : "Common",
          record.pricingMethod,
        ),
    pricingGroups: Array.isArray(snapshot.pricingGroups)
      ? snapshot.pricingGroups.map((group) => ({ ...group }))
      : Array.isArray(record.pricingGroups)
        ? record.pricingGroups.map((group) => ({ ...group }))
        : [],
    allocationRows: Array.isArray(snapshot.allocationRows)
      ? snapshot.allocationRows.map((row) => ({ ...row }))
      : Array.isArray(record.allocationRows)
        ? record.allocationRows.map((row) => ({ ...row }))
        : [],
    salesEmployeeName: snapshotHas("salesEmployeeName")
      ? snapshot.salesEmployeeName
      : pickValue(backendSite.incentiveEmployeeName, record.salesEmployeeName),
    companyId: snapshotHas("companyId")
      ? snapshot.companyId
      : pickValue(backendSite.incentiveEmployeeNumber, record.companyId),
    commercialTerms: Array.isArray(backendSite.commercialTerms)
      ? backendSite.commercialTerms.map((term) => ({ ...term }))
      : Array.isArray(record.commercialTerms)
        ? record.commercialTerms.map((term) => ({ ...term }))
        : [],
    stage1GroupRows: Array.isArray(snapshot.stage1GroupRows)
      ? snapshot.stage1GroupRows.map((row) => ({ ...row }))
      : Array.isArray(record.stage1GroupRows)
        ? record.stage1GroupRows.map((row) => ({ ...row }))
        : [],
  };
}

function hasExistingMouData(record) {
  return Boolean(
    normalizeValue(record?.mouSentDate) ||
      normalizeValue(record?.mouReceivedDate) ||
      normalizeValue(record?.mouStartDate) ||
      normalizeValue(record?.mouEndDate),
  );
}

function isCompletedSiteRow(row) {
  return Boolean(normalizeValue(row?.screenCode) && normalizeValue(row?.screenName));
}

function getApplicableSiteRows(rows = []) {
  const completedRows = rows.filter(isCompletedSiteRow);
  const hasSelectedRows = rows.some((row) => row?.selected);

  if (!hasSelectedRows) {
    return completedRows;
  }

  return completedRows.filter((row) => row.selected);
}

function normalizePersistedScreenRow(row = {}) {
  return {
    ...row,
    screenCode: normalizeIdInput(row.screenCode || row.screenId),
    screenName: normalizeValue(row.screenName || row.siteName),
  };
}

function getUniformFieldValue(rows = [], field) {
  const values = rows
    .map((row) => normalizeValue(row?.[field]))
    .filter(Boolean);

  if (values.length === 0) {
    return "";
  }

  const uniqueValues = Array.from(
    new Set(values.map((value) => value.toUpperCase())),
  );

  if (uniqueValues.length !== 1) {
    return null;
  }

  return values[0];
}

function getUniformIdFieldValue(rows = [], field, { allowStandaloneBlank = false } = {}) {
  const values = rows.map((row) =>
    normalizeIdInput(row?.[field], {
      allowStandaloneBlank,
    }),
  );

  const presentValues = values.filter(Boolean);

  if (presentValues.length === 0) {
    return "";
  }

  const uniqueValues = Array.from(new Set(presentValues));

  if (uniqueValues.length !== 1) {
    return null;
  }

  return uniqueValues[0];
}

function SiteEntry({
  onMoveToStage2 = () => {},
  readinessRecords = [],
  setReadinessRecords = () => {},
  existingStage2Records = [],
  existingStage3Records = [],
  setStage2Records = () => {},
  setStage3Records = () => {},
  isMinimized = false,
  isFullscreenActive = false,
  onRequestFullscreenRestore = () => {},
  onMinimizedChange = () => {},
  returnedStage1Record = null,
  onReturnedStage1RecordConsumed = () => {},
}) {
  const [siteCreationData, setSiteCreationData] = useState(null);
  const [mouSubscriptionData, setMouSubscriptionData] = useState(null);
  const [backendSiteRecords, setBackendSiteRecords] = useState([]);
  const [readinessSearchTerm, setReadinessSearchTerm] = useState("");
  const [stage1SaveNotice, setStage1SaveNotice] = useState("");
  const [stage2MoveNotice, setStage2MoveNotice] = useState("");
  const [isStage2MoveConfirmOpen, setIsStage2MoveConfirmOpen] = useState(false);
  const [isStage1BackendLoading, setIsStage1BackendLoading] = useState(true);
  const [isStage1BackendSaving, setIsStage1BackendSaving] = useState(false);
  const [siteCreationResetSignal, setSiteCreationResetSignal] = useState(0);
  const [mouSubscriptionResetSignal, setMouSubscriptionResetSignal] = useState(0);
  const [editingReadinessRecordId, setEditingReadinessRecordId] = useState("");
  const [editingReadinessSourceRecordId, setEditingReadinessSourceRecordId] = useState("");
  const [isEntryAreaCollapsed, setIsEntryAreaCollapsed] = useState(false);
  const [isReadinessCollapsed, setIsReadinessCollapsed] = useState(true);
  const [agreementChoice, setAgreementChoice] = useState("");
  const [isAgreementChoiceOpen, setIsAgreementChoiceOpen] = useState(false);
  const [correctionScreenCode, setCorrectionScreenCode] = useState("");
  const [addendumScope, setAddendumScope] = useState("");
  const [addendumSelectedSiteIds, setAddendumSelectedSiteIds] = useState([]);
  const readinessSelectAllRef = useRef(null);
  const siteCreationSectionRef = useRef(null);
  const mouSubscriptionSectionRef = useRef(null);
  const readinessSectionRef = useRef(null);
  const shouldRestoreBlankEntryOnExpandRef = useRef(false);
  const stage1SaveNoticeTimerRef = useRef(null);
  const stage2MoveNoticeTimerRef = useRef(null);
  const pendingStage2MoveRecordsRef = useRef([]);
  const previousMinimizedRef = useRef(isMinimized);

  const handleSiteCreationChange = useCallback((data) => {
    setSiteCreationData(data);
  }, []);

  const handleMouSubscriptionChange = useCallback((data) => {
    setMouSubscriptionData(data);
  }, []);

  useEffect(() => {
    let cancelled = false;

    async function loadStage1ReadinessFromBackend() {
      setIsStage1BackendLoading(true);

      try {
        const sites = await apiRequest("/sites");
        if (cancelled) return;

        setBackendSiteRecords(Array.isArray(sites) ? sites : []);

        const readyRecords = (Array.isArray(sites) ? sites : [])
          .filter((site) =>
            ["ready", "returned to stage 1"].includes(
              normalizeValue(site?.readinessStatus).toLowerCase(),
            ),
          )
          .map((site) => {
            const snapshot =
              site?.stage1Data && typeof site.stage1Data === "object"
                ? site.stage1Data
                : {};
            const currentCommercial = getLatestCommonCommercialTerm(site, sites);
            const useLatestCommonCommercial =
              site.foc !== true && currentCommercial?.pricingMode === "common";
            const snapshotHas = (key) =>
              Object.prototype.hasOwnProperty.call(snapshot, key);

            return {
              ...snapshot,
              backendSiteId: site.id,
              recordId: snapshot.recordId || site.id,
              billingCode: site.billingId || snapshot.billingCode || "",
              complexCode: site.complexId || snapshot.complexCode || "",
              screenCode: site.siteId || snapshot.screenCode || "",
              screenName: site.screenName || snapshot.screenName || "",
              location: site.location || snapshot.location || "",
              state: site.state || snapshot.state || "",
              siteType:
                snapshot.siteType ||
                (site.siteType === "Single" ? "Standalone" : "Complex"),
              foc: site.foc === true,
              processingStatus:
                site.processingStatus || snapshot.processingStatus || "New",
              readinessStatus:
                site.readinessStatus || snapshot.readinessStatus || "Ready",

              // 1B OTF / Subscription restore:
              // Prefer the exact Stage 1 UI snapshot. If the record predates the
              // snapshot field, rebuild the UI values from CommercialTerm.
              otfApplicable: snapshotHas("otfApplicable")
                ? snapshot.otfApplicable
                : currentCommercial?.otfApplicable || "No",
              otfType: snapshotHas("otfType")
                ? snapshot.otfType
                : mapOtfTypeFromApi(currentCommercial?.otfType),
              otfTaxMode: snapshotHas("otfTaxMode")
                ? snapshot.otfTaxMode
                : mapTaxModeFromApi(currentCommercial?.otfTaxMode),
              otfAmount: snapshotHas("otfAmount")
                ? snapshot.otfAmount
                : decimalToInputValue(currentCommercial?.otfAmount),
              subscriptionType: useLatestCommonCommercial
                ? mapSubscriptionTypeFromApi(currentCommercial?.subscriptionType)
                : snapshotHas("subscriptionType")
                ? snapshot.subscriptionType
                : mapSubscriptionTypeFromApi(
                    currentCommercial?.subscriptionType,
                  ),
              subscriptionMode: useLatestCommonCommercial
                ? mapSubscriptionModeFromApi(currentCommercial?.subscriptionMode)
                : snapshotHas("subscriptionMode")
                ? snapshot.subscriptionMode
                : mapSubscriptionModeFromApi(
                    currentCommercial?.subscriptionMode,
                  ),
              subscriptionFee: useLatestCommonCommercial
                ? decimalToInputValue(currentCommercial?.subscriptionFee)
                : snapshotHas("subscriptionFee")
                ? snapshot.subscriptionFee
                : decimalToInputValue(currentCommercial?.subscriptionFee),
              pricingMethod: snapshotHas("pricingMethod")
                ? snapshot.pricingMethod
                : currentCommercial?.pricingMode === "sitewise"
                  ? "Site-wise"
                  : currentCommercial?.pricingMode === "allocation"
                    ? "Allocation Model"
                    : "Common",
              pricingGroups: Array.isArray(snapshot.pricingGroups)
                ? snapshot.pricingGroups
                : [],
              allocationRows: Array.isArray(snapshot.allocationRows)
                ? snapshot.allocationRows
                : [],

              // Incentive Beneficiary restore:
              incentiveApplicable: snapshotHas("incentiveApplicable")
                ? snapshot.incentiveApplicable
                : site.incentiveApplicable === true,
              salesEmployeeName: snapshotHas("salesEmployeeName")
                ? snapshot.salesEmployeeName
                : site.incentiveEmployeeName || "",
              companyId: snapshotHas("companyId")
                ? snapshot.companyId
                : site.incentiveEmployeeNumber || "",

              selected: false,
              commercialTerms: site.commercialTerms || [],
            };
          });

        setReadinessRecords((currentRecords) => {
          // A Stage 2/Stage 3 -> Stage 1 return is shown optimistically before
          // its backend PATCH necessarily finishes. Do not let this initial
          // backend hydration erase that freshly returned row while the server
          // is catching up. Once the backend includes the row, the local
          // correction snapshot still wins so Edit opens with the exact handoff
          // data immediately.
          const returnedLocalRecords = (Array.isArray(currentRecords)
            ? currentRecords
            : []
          ).filter((record) => {
            const readiness = normalizeValue(record?.readinessStatus).toLowerCase();
            const backendRecord = (Array.isArray(sites) ? sites : []).find(
              (site) => site?.id === record?.backendSiteId,
            ) || (Array.isArray(sites) ? sites : []).find(
              (site) =>
                normalizeIdInput(site?.siteId) ===
                normalizeIdInput(record?.screenCode),
            );
            const backendIsStage1 = backendRecord
              ? ["ready", "returned to stage 1"].includes(
                  normalizeValue(backendRecord.readinessStatus).toLowerCase(),
                )
              : true;
            return (
              record?.stage1CorrectionReturn === true &&
              readiness === "ready" &&
              backendIsStage1 &&
              Boolean(normalizeIdInput(record?.screenCode))
            );
          });

          if (returnedLocalRecords.length === 0) {
            return readyRecords;
          }

          const returnedByScreenCode = new Map(
            returnedLocalRecords.map((record) => [
              normalizeIdInput(record.screenCode),
              record,
            ]),
          );

          const mergedBackendRecords = readyRecords.filter(
            (record) =>
              !returnedByScreenCode.has(normalizeIdInput(record?.screenCode)),
          );

          return [...returnedLocalRecords, ...mergedBackendRecords];
        });
      } catch (error) {
        if (!cancelled) {
          console.error("Unable to load Stage 1 Readiness:", error);
        }
      } finally {
        if (!cancelled) {
          setIsStage1BackendLoading(false);
        }
      }
    }

    loadStage1ReadinessFromBackend();

    return () => {
      cancelled = true;
    };
  }, [setReadinessRecords]);

  const handleSiteImportCompleted = useCallback(() => {
    window.requestAnimationFrame(() => {
      guidedScrollToElement(mouSubscriptionSectionRef.current);
    });
  }, []);


  const dismissStage1SaveNotice = useCallback(() => {
    if (stage1SaveNoticeTimerRef.current) {
      window.clearTimeout(stage1SaveNoticeTimerRef.current);
      stage1SaveNoticeTimerRef.current = null;
    }

    setStage1SaveNotice("");
  }, []);

  const showStage1SaveNotice = useCallback((message) => {
    if (stage1SaveNoticeTimerRef.current) {
      window.clearTimeout(stage1SaveNoticeTimerRef.current);
    }

    setStage1SaveNotice(message);
    stage1SaveNoticeTimerRef.current = window.setTimeout(() => {
      setStage1SaveNotice("");
      stage1SaveNoticeTimerRef.current = null;
    }, 5000);
  }, []);


  const dismissStage2MoveNotice = useCallback(() => {
    if (stage2MoveNoticeTimerRef.current) {
      window.clearTimeout(stage2MoveNoticeTimerRef.current);
      stage2MoveNoticeTimerRef.current = null;
    }

    setStage2MoveNotice("");
  }, []);

  const showStage2MoveNotice = useCallback((message) => {
    if (stage2MoveNoticeTimerRef.current) {
      window.clearTimeout(stage2MoveNoticeTimerRef.current);
    }

    setStage2MoveNotice(message);
    stage2MoveNoticeTimerRef.current = window.setTimeout(() => {
      setStage2MoveNotice("");
      stage2MoveNoticeTimerRef.current = null;
    }, 5000);
  }, []);

  function getBackendSiteForReadinessRecord(record) {
    const backendSiteId = normalizeValue(record?.backendSiteId);
    const screenCode = normalizeIdInput(record?.screenCode);

    return (
      backendSiteRecords.find((site) => normalizeValue(site?.id) === backendSiteId) ||
      backendSiteRecords.find(
        (site) => normalizeIdInput(site?.siteId) === screenCode,
      ) ||
      null
    );
  }

  function isReadinessRecordSelectable(record) {
    return isStage1ReadyForSelection(
      getBackendSiteForReadinessRecord(record) || record,
    );
  }

  const currentStage1ReadinessRecords = useMemo(() => {
    const backendBySiteId = new Map(
      backendSiteRecords
        .map((site) => [normalizeValue(site?.id), site])
        .filter(([siteId]) => Boolean(siteId)),
    );
    const backendByScreenCode = new Map(
      backendSiteRecords
        .map((site) => [normalizeIdInput(site?.siteId), site])
        .filter(([screenCode]) => Boolean(screenCode)),
    );

    const currentStageRecords = readinessRecords.filter((record) => {
      const backendSite =
        backendBySiteId.get(normalizeValue(record?.backendSiteId)) ||
        backendByScreenCode.get(normalizeIdInput(record?.screenCode));

      // Unsaved rows have no canonical backend record yet. Once a Site is
      // known, only its top-level backend stage may place it in this queue.
      if (!backendSite) {
        return true;
      }

      return getCurrentWorkflowStage(backendSite) === "Stage 1";
    });

    // A Stage 3 -> Stage 1 return can briefly coexist with its hydrated
    // backend row. The table is a projection of current Sites, so one
    // persisted Site ID must produce one readiness row. Unsaved rows use
    // their screen code only as a fallback identity.
    const uniqueRecords = new Map();
    currentStageRecords.forEach((record, index) => {
      const backendSite =
        backendBySiteId.get(normalizeValue(record?.backendSiteId)) ||
        backendByScreenCode.get(normalizeIdInput(record?.screenCode));
      const canonicalSiteId = normalizeValue(
        backendSite?.id || record?.backendSiteId,
      );
      const screenCode = normalizeIdInput(record?.screenCode);
      const identity = canonicalSiteId
        ? `backend:${canonicalSiteId}`
        : `local:${screenCode || normalizeValue(record?.recordId) || index}`;
      const existingRecord = uniqueRecords.get(identity);

      if (!existingRecord) {
        uniqueRecords.set(identity, record);
        return;
      }

      // Prefer the row carrying the canonical backend ID over an optimistic
      // local snapshot, while retaining the first row for genuinely local
      // records that have no backend identity yet.
      if (!existingRecord.backendSiteId && record.backendSiteId) {
        uniqueRecords.set(identity, record);
      }
    });

    return Array.from(uniqueRecords.values());
  }, [backendSiteRecords, readinessRecords]);

  const filteredReadinessRecords = useMemo(() => {
    const search = normalizeValue(readinessSearchTerm).toLowerCase();

    if (!search) {
      return currentStage1ReadinessRecords;
    }

    return currentStage1ReadinessRecords.filter((record) => {
      const searchableValues = [
        record.billingCode,
        record.complexCode,
        record.screenCode,
        record.screenName,
        record.location,
        record.readinessStatus,
        record.status,
      ];

      return searchableValues.some((value) =>
        normalizeValue(value).toLowerCase().includes(search),
      );
    });
  }, [currentStage1ReadinessRecords, readinessSearchTerm]);

  const hasSelectedReadinessRecords = currentStage1ReadinessRecords.some(
    (record) => record.selected && isReadinessRecordSelectable(record),
  );
  const selectableVisibleReadinessRecords = filteredReadinessRecords.filter(
    (record) => isReadinessRecordSelectable(record),
  );
  const hasVisibleReadinessRecords = selectableVisibleReadinessRecords.length > 0;
  const areAllVisibleReadinessSelected =
    hasVisibleReadinessRecords &&
    selectableVisibleReadinessRecords.every((record) => record.selected);
  const areSomeVisibleReadinessSelected =
    hasVisibleReadinessRecords &&
    selectableVisibleReadinessRecords.some((record) => record.selected) &&
    !areAllVisibleReadinessSelected;

  useEffect(() => {
    if (readinessSelectAllRef.current) {
      readinessSelectAllRef.current.indeterminate =
        areSomeVisibleReadinessSelected;
    }
  }, [areSomeVisibleReadinessSelected]);

  useEffect(() => {
    return () => {
      if (stage1SaveNoticeTimerRef.current) {
        window.clearTimeout(stage1SaveNoticeTimerRef.current);
      }
      if (stage2MoveNoticeTimerRef.current) {
        window.clearTimeout(stage2MoveNoticeTimerRef.current);
      }
    };
  }, []);

  useEffect(() => {
    const wasMinimized = previousMinimizedRef.current;
    previousMinimizedRef.current = isMinimized;

    if (wasMinimized && !isMinimized && shouldRestoreBlankEntryOnExpandRef.current) {
      setIsEntryAreaCollapsed(false);
      setIsReadinessCollapsed(true);
      shouldRestoreBlankEntryOnExpandRef.current = false;
    }
  }, [isMinimized]);

  useEffect(() => {
    if (!returnedStage1Record) {
      return;
    }

    const renewalGroupRecords =
      returnedStage1Record?.complexRenewal === true &&
      Array.isArray(returnedStage1Record?.renewalGroupRecords)
        ? returnedStage1Record.renewalGroupRecords
        : [];

    const incomingRecords =
      renewalGroupRecords.length > 0
        ? renewalGroupRecords
        : [returnedStage1Record];

    const returnedRecords = incomingRecords.map((record) => ({
        ...record,
        selected: false,
        readinessStatus: record.readinessStatus || "Ready",
        stage1CorrectionReturn: true,
      }));

    setReadinessRecords((currentRecords) =>
      upsertReadinessRecords(currentRecords, returnedRecords),
    );

    // Returned/existing sites follow the same untouched Stage 1 Readiness
    // behavior as new sites. Do not auto-open Site Entry or edit mode.
    setEditingReadinessRecordId("");
    setEditingReadinessSourceRecordId("");
    setCorrectionScreenCode("");
    setAgreementChoice("");
    setAddendumScope("");
    setAddendumSelectedSiteIds([]);
    setIsAgreementChoiceOpen(false);
    setSiteCreationData(null);
    setMouSubscriptionData(null);
    setIsEntryAreaCollapsed(true);
    onMinimizedChange(true);
    setIsReadinessCollapsed(false);
    setReadinessSearchTerm("");

    window.requestAnimationFrame(() => {
      guidedScrollToElement(readinessSectionRef.current);
    });

    onReturnedStage1RecordConsumed();
  }, [onReturnedStage1RecordConsumed, returnedStage1Record]);

  const editingReadinessGroupRecords = useMemo(() => {
    if (!editingReadinessRecordId) {
      return [];
    }

    const groupRecords = readinessRecords.filter(
      (record) => getStage1GroupId(record) === editingReadinessRecordId,
    );

    // Keep the screen the user actually clicked as the first edit source.
    // MouSubscription may still fall back to a non-FoC sibling when the
    // clicked screen itself is FoC, but it must not silently use another
    // billable sibling when the clicked screen is already non-FoC.
    if (!editingReadinessSourceRecordId) {
      return groupRecords;
    }

    return [...groupRecords].sort((left, right) => {
      if (left.recordId === editingReadinessSourceRecordId) return -1;
      if (right.recordId === editingReadinessSourceRecordId) return 1;
      return 0;
    });
  }, [
    editingReadinessRecordId,
    editingReadinessSourceRecordId,
    readinessRecords,
  ]);

  const editingReadinessCanonicalRecords = useMemo(
    () =>
      editingReadinessGroupRecords
        .map((record) =>
          buildCanonicalStage1EditRecord(record, backendSiteRecords),
        )
        .filter(Boolean),
    [backendSiteRecords, editingReadinessGroupRecords],
  );

  const editingReadinessGroupRows = useMemo(() => {
    const sourceGroupRecord = editingReadinessCanonicalRecords[0];
    const sourceGroupId = getStage1GroupId(sourceGroupRecord);
    const sourceBillingCode = normalizeIdInput(sourceGroupRecord?.billingCode);
    const sourceComplexCode = normalizeIdInput(sourceGroupRecord?.complexCode, {
      allowStandaloneBlank: true,
    });
    const persistedSnapshotRows = backendSiteRecords
      .filter((site) => {
        const snapshot =
          site?.stage1Data && typeof site.stage1Data === "object"
            ? site.stage1Data
            : {};
        const snapshotGroupId = getStage1GroupId(snapshot);
        const sameGroup = sourceGroupId && snapshotGroupId === sourceGroupId;
        const sameBusinessGroup =
          normalizeIdInput(site?.billingId) === sourceBillingCode &&
          normalizeIdInput(site?.complexId, { allowStandaloneBlank: true }) ===
            sourceComplexCode;

        // Group snapshots can carry different historical record IDs after a
        // return/correction. The billing + complex identity is the stable
        // membership key for this Stage 1 edit projection.
        return sameGroup || sameBusinessGroup;
      })
      .flatMap((site) => {
        const snapshot =
          site?.stage1Data && typeof site.stage1Data === "object"
            ? site.stage1Data
            : {};
        return Array.isArray(snapshot.stage1GroupRows)
          ? snapshot.stage1GroupRows
          : [];
      });
    const snapshotRows = [
      ...(Array.isArray(sourceGroupRecord?.stage1GroupRows)
        ? sourceGroupRecord.stage1GroupRows
        : []),
      ...persistedSnapshotRows,
    ].reduce((rows, row) => {
      const screenCode = normalizeIdInput(row?.screenCode);
      const backendScopeSite = backendSiteRecords.find(
        (site) => normalizeIdInput(site?.siteId) === screenCode,
      );
      if (
        !screenCode ||
        row?.foc === true ||
        backendScopeSite?.foc === true ||
        rows.some((item) => normalizeIdInput(item?.screenCode) === screenCode)
      ) {
        return rows;
      }
      rows.push(row);
      return rows;
    }, []);

    if (snapshotRows.length > 0) {
      return snapshotRows.map((row) => ({
        ...row,
        selected: correctionScreenCode
          ? normalizeIdInput(row?.screenCode) === correctionScreenCode
          : false,
      }));
    }

    return editingReadinessCanonicalRecords.map((record) => ({
      billingCode: record.billingCode || "",
      complexCode: record.complexCode || "",
      billingName: record.billingName || "",
      screenCode: record.screenCode || "",
      screenName: record.screenName || "",
      location: record.location || "",
      state: record.state || "",
      siteType:
        record.siteType || (normalizeValue(record.complexCode) ? "Complex" : "Standalone"),
      selected: false,
      status: record.status || "Complete",
    }));
  }, [backendSiteRecords, correctionScreenCode, editingReadinessCanonicalRecords]);

  const blockedStage1SiteIds = useMemo(() => {
    const blockedIds = new Set();

    // Stage 1 rows that belong to the group currently being edited are allowed
    // to stay editable. Unrelated Stage 1 rows remain blocked.
    readinessRecords.forEach((record) => {
      const screenCode = normalizeIdInput(record?.screenCode);

      if (!screenCode) {
        return;
      }

      if (
        correctionScreenCode &&
        screenCode !== correctionScreenCode
      ) {
        blockedIds.add(screenCode);
        return;
      }

      if (
        editingReadinessRecordId &&
        getStage1GroupId(record) === editingReadinessRecordId
      ) {
        return;
      }

      blockedIds.add(screenCode);
    });

    // Downstream-active screens must never be re-created in Stage 1 merely
    // because they share the same Stage 1 group snapshot. Explicit correction,
    // Addendum, or Renewal routes remove the applicable screen(s) from the
    // downstream stage before handing them back to Stage 1.
    [...existingStage2Records, ...existingStage3Records].forEach((record) => {
      const screenCode = normalizeIdInput(record?.screenCode);

      if (screenCode) {
        blockedIds.add(screenCode);
      }
    });

    // Also lock sibling rows that exist only inside the preserved complex snapshot.
    if (correctionScreenCode) {
      editingReadinessGroupRows.forEach((row) => {
        const screenCode = normalizeIdInput(row?.screenCode);
        if (screenCode && screenCode !== correctionScreenCode) {
          blockedIds.add(screenCode);
        }
      });
    }

    return Array.from(blockedIds);
  }, [
    correctionScreenCode,
    editingReadinessGroupRows,
    editingReadinessRecordId,
    existingStage2Records,
    existingStage3Records,
    readinessRecords,
  ]);

  const addendumAvailable = useMemo(() => {
    const applicableRows = getApplicableSiteRows(siteCreationData?.siteRows || []).filter(
      (row) =>
        !blockedStage1SiteIds.includes(normalizeIdInput(row?.screenCode)),
    );

    if (applicableRows.length === 0) {
      return false;
    }

    const billingCode =
      normalizeIdInput(siteCreationData?.billingCode) ||
      getUniformIdFieldValue(applicableRows, "billingCode");

    if (!billingCode) {
      return false;
    }

    const selectedSiteIds = new Set(
      applicableRows
        .map((row) => normalizeIdInput(row.screenCode))
        .filter(Boolean),
    );

    const historyRecords = [
      ...readinessRecords.map((record) => ({ record, downstream: false })),
      ...existingStage2Records.map((record) => ({ record, downstream: true })),
      ...existingStage3Records.map((record) => ({ record, downstream: true })),
    ];

    return historyRecords.some(({ record, downstream }) => {
      if (!downstream && !hasExistingMouData(record)) {
        return false;
      }

      if (
        editingReadinessRecordId &&
        getStage1GroupId(record) === editingReadinessRecordId
      ) {
        return false;
      }

      const recordBillingId = normalizeIdInput(record.billingCode);
      if (recordBillingId !== billingCode) {
        return false;
      }

      const recordSiteId = normalizeIdInput(record.screenCode);
      return !recordSiteId || !selectedSiteIds.has(recordSiteId);
    });
  }, [
    blockedStage1SiteIds,
    editingReadinessRecordId,
    existingStage2Records,
    existingStage3Records,
    readinessRecords,
    siteCreationData,
  ]);

  const billingCodeScopeRecords = useMemo(() => {
    const directBillingCode = normalizeIdInput(siteCreationData?.billingCode);
    const applicableRows = getApplicableSiteRows(
      siteCreationData?.siteRows || [],
    );
    const billingCode =
      directBillingCode ||
      getUniformIdFieldValue(applicableRows, "billingCode");

    if (!billingCode) {
      return [];
    }

    const allRecords = [
      ...readinessRecords,
      ...existingStage2Records,
      ...existingStage3Records,
    ];

    const bySite = new Map();

    function addScopeSite(candidate, fallbackRecord = null) {
      const candidateBillingCode =
        normalizeIdInput(candidate?.billingCode) ||
        normalizeIdInput(fallbackRecord?.billingCode);

      if (candidateBillingCode !== billingCode) {
        return;
      }

      const siteId = normalizeIdInput(candidate?.screenCode);
      if (!siteId) {
        return;
      }

      const existing = bySite.get(siteId);

      // Prefer the live workflow record when one exists, but preserve snapshot
      // values as fallback so the Addendum selector can still show every site
      // under the Billing Code even when only one site is currently in Stage 1.
      const nextRecord = {
        ...(candidate || {}),
        ...(fallbackRecord || {}),
        billingCode,
        screenCode: siteId,
        screenName:
          normalizeValue(candidate?.screenName) ||
          normalizeValue(fallbackRecord?.screenName),
        complexCode:
          normalizeIdInput(candidate?.complexCode, {
            allowStandaloneBlank: true,
          }) ||
          normalizeIdInput(fallbackRecord?.complexCode, {
            allowStandaloneBlank: true,
          }),
      };

      if (!existing) {
        bySite.set(siteId, nextRecord);
        return;
      }

      bySite.set(siteId, {
        ...nextRecord,
        ...existing,
        screenCode: siteId,
        billingCode,
        screenName:
          normalizeValue(existing?.screenName) ||
          normalizeValue(nextRecord?.screenName),
        complexCode:
          normalizeIdInput(existing?.complexCode, {
            allowStandaloneBlank: true,
          }) ||
          normalizeIdInput(nextRecord?.complexCode, {
            allowStandaloneBlank: true,
          }),
      });
    }

    allRecords.forEach((record) => {
      if (normalizeIdInput(record?.billingCode) !== billingCode) {
        return;
      }

      // Add the live Stage 1 / Stage 2 / Stage 3 site.
      addScopeSite(record);

      // Also expand the preserved Stage 1 group snapshot. This is important
      // when a single site is returned for correction: sibling sites can remain
      // downstream but must still appear in the Addendum scope selector.
      const snapshotRows = Array.isArray(record?.stage1GroupRows)
        ? record.stage1GroupRows
        : [];

      snapshotRows.forEach((row) => {
        addScopeSite(row, record);
      });
    });

    // Include the currently loaded Site Creation rows as well. This keeps newly
    // introduced / edited sites visible in the same Billing Code scope before save.
    applicableRows.forEach((row) => {
      addScopeSite(row, {
        billingCode,
        complexCode: siteCreationData?.complexCode || "",
        billingName: siteCreationData?.billingName || "",
        location: siteCreationData?.location || "",
        state: siteCreationData?.state || "",
      });
    });

    return Array.from(bySite.values()).sort((left, right) =>
      normalizeIdInput(left?.screenCode).localeCompare(
        normalizeIdInput(right?.screenCode),
        undefined,
        { numeric: true },
      ),
    );
  }, [
    existingStage2Records,
    existingStage3Records,
    readinessRecords,
    siteCreationData,
  ]);

  const eligibleSiteRowsForStage1 = useMemo(
    () =>
      getApplicableSiteRows(siteCreationData?.siteRows || []).filter((row) => {
        const screenCode = normalizeIdInput(row?.screenCode);

        if (correctionScreenCode && screenCode !== correctionScreenCode) {
          return false;
        }

        return !blockedStage1SiteIds.includes(screenCode);
      }),
    [blockedStage1SiteIds, correctionScreenCode, siteCreationData],
  );

  const requiredAgreementType = addendumAvailable ? "Addendum" : "MoU";

  useEffect(() => {
    if (editingReadinessRecordId) {
      return;
    }

    setAgreementChoice("");
    setIsAgreementChoiceOpen(false);
  }, [
    editingReadinessRecordId,
    siteCreationData?.billingCode,
    siteCreationData?.complexCode,
  ]);

  function handleRequestAgreementChoice() {
    if (!validateSiteCreation()) {
      return;
    }

    if (eligibleSiteRowsForStage1.length === 0) {
      alert("No eligible new Screen Codes are available for MoU / Addendum processing.");
      return;
    }

    setAgreementChoice(requiredAgreementType);
    if (requiredAgreementType === "Addendum") {
      setAddendumScope("Billing Code");
      setAddendumSelectedSiteIds(
        billingCodeScopeRecords.map((record) => normalizeIdInput(record.screenCode)),
      );
    }
    setIsAgreementChoiceOpen(true);
  }

  function handleConfirmAgreementChoice() {
    if (!agreementChoice) {
      return;
    }

    if (agreementChoice === "Addendum") {
      if (!addendumScope) {
        alert("Select the Addendum scope.");
        return;
      }
      if (addendumScope === "Selected Sites" && addendumSelectedSiteIds.length === 0) {
        alert("Select at least one site for the Addendum.");
        return;
      }
    }

    setIsAgreementChoiceOpen(false);
  }

  function validateSiteCreation() {
    if (!siteCreationData) {
      alert("Complete the Site Creation details.");
      return false;
    }

    const {
      billingCode,
      complexCode,
      location,
      state,
      processingStatus,
      siteRows = [],
    } = siteCreationData;
    const editRows = editingReadinessGroupRows.map(normalizePersistedScreenRow);
    const sourceRows =
      editingReadinessRecordId && editRows.length > 0 ? editRows : siteRows;
    const completedRows = getApplicableSiteRows(sourceRows).filter((row) => {
      const screenCode = normalizeIdInput(row?.screenCode);

      if (correctionScreenCode && screenCode !== correctionScreenCode) {
        return false;
      }

      return !blockedStage1SiteIds.includes(screenCode);
    });
    const normalizedBillingId = normalizeIdInput(billingCode);
    const resolvedBillingId =
      normalizedBillingId ||
      getUniformIdFieldValue(completedRows, "billingCode");
    const resolvedLocation = normalizeValue(location) || getUniformFieldValue(completedRows, "location");
    const resolvedState = normalizeValue(state) || getUniformFieldValue(completedRows, "state");

    const invalidBillingRow = completedRows.find((row) => {
      const normalizedRowBillingId = normalizeIdInput(row.billingCode);
      return normalizedRowBillingId && !isValidIdInput(normalizedRowBillingId);
    });
    if (normalizedBillingId && !isValidIdInput(normalizedBillingId)) {
      alert(getIdValidationMessage("Billing Code / Customer Code"));
      return false;
    }

    if (invalidBillingRow) {
      alert(getIdValidationMessage("Billing Code / Customer Code"));
      return false;
    }

    const invalidSiteRow = completedRows.find(
      (row) => !isValidIdInput(normalizeValue(row.screenCode)),
    );
    if (invalidSiteRow) {
      alert(getIdValidationMessage("Screen Code"));
      return false;
    }

    if (!resolvedBillingId) {
      alert(getIdValidationMessage("Billing Code / Customer Code"));
      return false;
    }

    if (!resolvedLocation) {
      alert("Enter the Location.");
      return false;
    }

    if (!resolvedState) {
      alert("Enter the State.");
      return false;
    }

    if (!processingStatus) {
      alert("Select Processing Status.");
      return false;
    }

    if (completedRows.length === 0) {
      alert("Add at least one Screen Code and Screen Name.");
      return false;
    }

    const siteIds = completedRows.map((row) =>
      row.screenCode.trim().toUpperCase(),
    );

    const duplicateSiteIds = siteIds.filter(
      (screenCode, index) => siteIds.indexOf(screenCode) !== index,
    );

    if (duplicateSiteIds.length > 0) {
      alert(`Duplicate Screen Code found: ${duplicateSiteIds[0]}`);
      return false;
    }

    return true;
  }

  async function handleSaveStage1() {
    if (isStage1BackendSaving) {
      return;
    }

    if (!validateSiteCreation()) {
      return;
    }

    if (!mouSubscriptionData) {
      alert("Complete the Subscription details.");
      return;
    }

    const editRows = editingReadinessGroupRows.map(normalizePersistedScreenRow);
    const sourceRows =
      editingReadinessRecordId && editRows.length > 0
        ? editRows
        : siteCreationData.siteRows;
    const completedRows = getApplicableSiteRows(sourceRows).filter((row) => {
      const screenCode = normalizeIdInput(row?.screenCode);

      if (correctionScreenCode && screenCode !== correctionScreenCode) {
        return false;
      }

      return !blockedStage1SiteIds.includes(screenCode);
    });
    const resolvedBillingId =
      normalizeIdInput(siteCreationData.billingCode) ||
      getUniformIdFieldValue(completedRows, "billingCode");
    const resolvedLocation = normalizeValue(siteCreationData.location) || getUniformFieldValue(completedRows, "location");
    const resolvedState = normalizeValue(siteCreationData.state) || getUniformFieldValue(completedRows, "state");
    const resolvedComplexId = normalizeIdInput(siteCreationData.complexCode, {
      allowStandaloneBlank: true,
    });
    const billingIdValue = resolvedBillingId || "";
    const locationValue = resolvedLocation || "";
    const stateValue = resolvedState || "";
    const complexIdValue = resolvedComplexId || "";

    const existingSiteIds = new Set(
      readinessRecords
        .filter((record) => getStage1GroupId(record) !== editingReadinessRecordId)
        .map((record) => normalizeIdInput(record.screenCode)),
    );

    const duplicateExistingSite = completedRows.find((row) =>
      existingSiteIds.has(normalizeIdInput(row.screenCode)),
    );

    if (duplicateExistingSite) {
      alert(
        `${duplicateExistingSite.screenCode} already exists in Stage 1 Readiness.`,
      );
      return;
    }

    const savedAt = new Date().toISOString();
    const stage1GroupId = editingReadinessRecordId || crypto.randomUUID();
    const commonSelectedSiteChange =
      !completedRows.some((row) => row.foc === true) &&
      normalizeValue(mouSubscriptionData.pricingMethod).toLowerCase() ===
        "common" &&
      Boolean(complexIdValue);
    const persistedCommonScopeRows = backendSiteRecords
      .filter(
        (site) =>
          normalizeIdInput(site?.billingId) === billingIdValue &&
          normalizeIdInput(site?.complexId, { allowStandaloneBlank: true }) ===
            complexIdValue &&
          site?.foc !== true,
      )
      .flatMap((site) => {
        const snapshot =
          site?.stage1Data && typeof site.stage1Data === "object"
            ? site.stage1Data
            : {};
        return Array.isArray(snapshot.stage1GroupRows)
          ? snapshot.stage1GroupRows
          : [];
      });
    const commonScopeRows = [
      ...persistedCommonScopeRows,
      ...completedRows,
    ].reduce((rows, row) => {
      const screenCode = normalizeIdInput(row?.screenCode);
      if (
        !screenCode ||
        rows.some((item) => normalizeIdInput(item?.screenCode) === screenCode)
      ) {
        return rows;
      }
      rows.push({
        ...row,
        screenCode,
        selected: false,
        status: row.status || "Complete",
      });
      return rows;
    }, []);
    if (commonSelectedSiteChange && commonScopeRows.length > 0) {
      const commonScopeScreenCodes = new Set(
        commonScopeRows.map((row) => normalizeIdInput(row?.screenCode)),
      );
      const billingStartedScreen = backendSiteRecords.find(
        (site) =>
          commonScopeScreenCodes.has(normalizeIdInput(site?.siteId)) &&
          site?.foc !== true &&
          hasBillingActuallyStarted(site),
      );

      if (billingStartedScreen) {
        alert(
          `Common pricing for ${billingStartedScreen.siteId} cannot be changed through 1B after billing has started. Use Price Change & Reallocation once for the affected Common scope.`,
        );
        return;
      }
    }
    const stage1GroupRows =
      commonSelectedSiteChange && commonScopeRows.length > 0
        ? commonScopeRows
        : correctionScreenCode && editingReadinessGroupRows.length > 0
          ? editingReadinessGroupRows.map((row) => ({
              ...row,
              selected: false,
              status: row.status || "Complete",
            }))
          : completedRows.map((row) => ({
              ...row,
              siteType:
                row.siteType ||
                (isStandaloneComplexCode(row.complexCode)
                  ? "Standalone"
                  : "Complex"),
              selected: false,
              status: row.status || "Complete",
            }));

    const newRecords = completedRows.map((row) => {
      const normalizedScreenCode = normalizeIdInput(row.screenCode);
      const correctionSourceRecord =
        editingReadinessGroupRecords.find(
          (record) =>
            normalizeIdInput(record?.screenCode) === normalizedScreenCode,
        ) ||
        (editingReadinessGroupRecords.length === 1
          ? editingReadinessGroupRecords[0]
          : null);
      const isCorrectionReturn = Boolean(
        correctionSourceRecord?.stage1CorrectionReturn,
      );

      return {
        ...(isCorrectionReturn ? correctionSourceRecord : {}),
        recordId:
          isCorrectionReturn && correctionSourceRecord?.recordId
            ? correctionSourceRecord.recordId
            : crypto.randomUUID(),
        stage1GroupId,
        stage1GroupRows,
        screenCode: normalizedScreenCode,
        screenName: row.screenName.trim(),
        billingCode: normalizeIdInput(billingIdValue),
        complexCode: normalizeIdInput(complexIdValue, {
          allowStandaloneBlank: true,
        }),
        billingName: normalizeValue(siteCreationData.billingName),
        location: locationValue.trim(),
        state: stateValue.trim(),
        siteType:
          row.siteType ||
          (isStandaloneComplexCode(row.complexCode)
            ? "Standalone"
            : "Complex"),
        processingStatus: isCorrectionReturn
          ? "Stage 1"
          : siteCreationData.processingStatus || "New",
        foc: Boolean(row.foc),
        commercialApplicable: !Boolean(row.foc),
        incentiveApplicable:
          !Boolean(row.foc) &&
          mouSubscriptionData.incentiveApplicable !== "Not Applicable",
        mouStatus: mouSubscriptionData.mouStatus,
        mouSentDate: mouSubscriptionData.mouSentDate,
        mouReceivedDate: mouSubscriptionData.mouReceivedDate,
        mouStartDate: mouSubscriptionData.mouStartDate,
        mouEndDate: mouSubscriptionData.mouEndDate,
        otfApplicable: row.foc ? "No" : mouSubscriptionData.otfApplicable,
        otfType: row.foc ? "" : mouSubscriptionData.otfType,
        otfTaxMode: row.foc
          ? "Not Applicable"
          : mouSubscriptionData.otfTaxMode,
        otfAmount: row.foc ? "" : mouSubscriptionData.otfAmount,
        subscriptionType: row.foc ? "Not Applicable" : mouSubscriptionData.subscriptionType,
        subscriptionMode: row.foc ? "" : mouSubscriptionData.subscriptionMode,
        subscriptionFee: row.foc ? "" : mouSubscriptionData.subscriptionFee,
        extensionRequired: mouSubscriptionData.extensionRequired,
        extensionRenewal: mouSubscriptionData.extensionRenewal,
        newMouStartDate: mouSubscriptionData.newMouStartDate,
        newMouEndDate: mouSubscriptionData.newMouEndDate,
        extensionRemarks: mouSubscriptionData.extensionRemarks,
        salesEmployeeName:
          row.foc || mouSubscriptionData.incentiveApplicable === "Not Applicable"
            ? "Not Applicable"
            : mouSubscriptionData.salesEmployeeName.trim(),
        companyId:
          row.foc || mouSubscriptionData.incentiveApplicable === "Not Applicable"
            ? ""
            : mouSubscriptionData.companyId,
        pricingMethod: row.foc ? "FoC" : mouSubscriptionData.pricingMethod,
        pricingGroups: Array.isArray(mouSubscriptionData.pricingGroups)
          ? mouSubscriptionData.pricingGroups.map((group) => ({ ...group }))
          : [],
        allocationRows: Array.isArray(mouSubscriptionData.allocationRows)
          ? mouSubscriptionData.allocationRows.map((allocationRow) => ({
              ...allocationRow,
            }))
          : [],
        allocationSettingsSnapshot:
          mouSubscriptionData.allocationSettingsSnapshot || null,
        agreementType:
          agreementChoice ||
          mouSubscriptionData.agreementType ||
          requiredAgreementType,
        addendumScope:
          (agreementChoice || mouSubscriptionData.agreementType || requiredAgreementType) === "Addendum"
            ? addendumScope || "Selected Sites"
            : "",
        addendumApplicableSiteIds:
          (agreementChoice || mouSubscriptionData.agreementType || requiredAgreementType) === "Addendum"
            ? (addendumScope === "Billing Code"
                ? billingCodeScopeRecords.map((record) => normalizeIdInput(record.screenCode))
                : addendumSelectedSiteIds)
            : [],
        readinessStatus: "Ready",
        stage1CorrectionReturn: isCorrectionReturn,
        selected: false,
        savedAt,
      };
    });

    setIsStage1BackendSaving(true);

    let persistedRecords;
    const persistedCommercialSiteIds = new Set();
    const persistedCommercialTerms = [];

    try {
      persistedRecords = [];

      for (const record of newRecords) {
        const recordBackendSiteId = normalizeValue(record?.backendSiteId);
        const recordScreenCode = normalizeIdInput(record?.screenCode);
        const existingBackendById = backendSiteRecords.find(
          (site) => normalizeValue(site?.id) === recordBackendSiteId,
        );
        const existingBackendByScreenCode = backendSiteRecords.find(
          (site) => normalizeIdInput(site?.siteId) === recordScreenCode,
        );

        if (
          existingBackendById &&
          normalizeIdInput(existingBackendById?.siteId) !== recordScreenCode
        ) {
          throw new Error(
            `Screen identity mismatch for ${recordScreenCode}. The saved Site ID belongs to ${normalizeIdInput(existingBackendById?.siteId)}.`,
          );
        }

        const existingBackendRecord =
          existingBackendByScreenCode ||
          existingBackendById ||
          readinessRecords.find(
            (currentRecord) =>
              normalizeValue(currentRecord?.backendSiteId) === recordBackendSiteId ||
              normalizeIdInput(currentRecord?.screenCode) === recordScreenCode,
          ) ||
          {};
        const readinessScreenCode = normalizeIdInput(existingBackendRecord?.screenCode);

        if (readinessScreenCode && readinessScreenCode !== recordScreenCode) {
          throw new Error(
            `Screen identity mismatch for ${recordScreenCode}. A readiness record for ${readinessScreenCode} cannot be updated.`,
          );
        }

        const persistedBackendSiteId =
          normalizeValue(existingBackendByScreenCode?.id) ||
          recordBackendSiteId ||
          normalizeValue(existingBackendRecord?.id) ||
          normalizeValue(existingBackendRecord?.backendSiteId);

        const sitePayload = {
          billingId: record.billingCode,
          complexId: record.complexCode || null,
          siteId: record.screenCode,
          screenName: record.screenName,
          location: record.location,
          state: record.state,
          siteType: mapSiteTypeForApi(record.siteType),
          foc: record.foc === true,
          processingStatus: record.processingStatus || "New",
          readinessStatus: "Ready",
          incentiveApplicable: record.incentiveApplicable === true,
          incentiveEmployeeName:
            record.incentiveApplicable === true
              ? record.salesEmployeeName || null
              : null,
          incentiveEmployeeNumber:
            record.incentiveApplicable === true
              ? record.companyId || null
              : null,
          stage1Data: {
            ...record,
            selected: false,
            readinessStatus: "Ready",
          },
        };

        let savedSite;

        if (persistedBackendSiteId) {
          savedSite = await apiRequest(
            `/sites/${persistedBackendSiteId}`,
            {
              method: "PATCH",
              body: JSON.stringify(sitePayload),
            },
          );
        } else {
          savedSite = await apiRequest("/sites", {
            method: "POST",
            body: JSON.stringify(sitePayload),
          });
        }

        if (!record.foc) {
          const savedCommercialTerm = await apiRequest("/commercial", {
            method: "POST",
            body: JSON.stringify({
              siteId: savedSite.id,
              pricingMode: mapPricingModeForApi(record.pricingMethod),
              otfApplicable: record.otfApplicable || "No",
              otfType:
                record.otfApplicable === "Yes"
                  ? record.otfType || null
                  : null,
              otfAmount:
                record.otfApplicable === "Yes"
                  ? toNumberOrNull(record.otfAmount)
                  : null,
              otfTaxMode:
                record.otfApplicable === "Yes"
                  ? record.otfTaxMode || null
                  : null,
              subscriptionType: mapSubscriptionTypeForApi(
                record.subscriptionType,
              ),
              subscriptionMode:
                record.subscriptionMode || null,
              subscriptionFee: toNumberOrNull(record.subscriptionFee),
            }),
          });
          if (savedCommercialTerm) {
            persistedCommercialTerms.push(savedCommercialTerm);
          }
          persistedCommercialSiteIds.add(savedSite.id);
        }

        if (record.correctionAuditId) {
          const correctionFields = Array.isArray(record.correctionFields)
            ? record.correctionFields
            : [];
          const newValues = correctionFields.reduce((values, field) => {
            values[field] = record[field] ?? null;
            return values;
          }, {});

          await apiRequest(`/sites/${savedSite.id}/corrections`, {
            method: "POST",
            body: JSON.stringify({
              auditId: record.correctionAuditId,
              eventId: record.correctionEventId,
              workflowType: "RECORD_CORRECTION",
              correctionFields,
              reason: record.correctionReason || "Record Correction",
              newValues,
            }),
          });
        }

        persistedRecords.push({
          ...record,
          backendSiteId: savedSite.id,
          readinessStatus: "Ready",
          selected: false,
        });
      }

      if (commonSelectedSiteChange) {
        const commonScopeScreenCodes = new Set(
          stage1GroupRows.map((row) => normalizeIdInput(row?.screenCode)),
        );
        const commonScopeSites = backendSiteRecords.filter(
          (site) =>
            commonScopeScreenCodes.has(normalizeIdInput(site?.siteId)) &&
            site?.foc !== true &&
            normalizeIdInput(site?.billingId) === billingIdValue &&
            normalizeIdInput(site?.complexId, { allowStandaloneBlank: true }) ===
              complexIdValue &&
            site?.id &&
            !persistedCommercialSiteIds.has(site.id),
        );

        for (const site of commonScopeSites) {
          const savedCommercialTerm = await apiRequest("/commercial", {
            method: "POST",
            body: JSON.stringify({
              siteId: site.id,
              pricingMode: mapPricingModeForApi(mouSubscriptionData.pricingMethod),
              otfApplicable: mouSubscriptionData.otfApplicable || "No",
              otfType:
                mouSubscriptionData.otfApplicable === "Yes"
                  ? mouSubscriptionData.otfType || null
                  : null,
              otfAmount:
                mouSubscriptionData.otfApplicable === "Yes"
                  ? toNumberOrNull(mouSubscriptionData.otfAmount)
                  : null,
              otfTaxMode:
                mouSubscriptionData.otfApplicable === "Yes"
                  ? mouSubscriptionData.otfTaxMode || null
                  : null,
              subscriptionType: mapSubscriptionTypeForApi(
                mouSubscriptionData.subscriptionType,
              ),
              subscriptionMode: mouSubscriptionData.subscriptionMode || null,
              subscriptionFee: toNumberOrNull(mouSubscriptionData.subscriptionFee),
            }),
          });
          if (savedCommercialTerm) {
            persistedCommercialTerms.push(savedCommercialTerm);
          }
        }
      }

    } catch (error) {
      console.error("Unable to persist Stage 1:", error);
      alert(
        `Stage 1 could not be saved to the backend. ${error.message}`,
      );
      return;
    } finally {
      setIsStage1BackendSaving(false);
    }

    const commercialTermsBySiteId = new Map(
      persistedCommercialTerms
        .filter((term) => normalizeValue(term?.siteId))
        .map((term) => [normalizeValue(term.siteId), term]),
    );
    const reconcileCommercialState = (records) =>
      records.map((record) => {
        const term = commercialTermsBySiteId.get(
          normalizeValue(record?.backendSiteId),
        );
        if (!term) return record;

        const completed =
          normalizeValue(record?.firstTimeBillingValidationStatus).toLowerCase() ===
            "completed" ||
          normalizeValue(record?.billingVerificationStatus).toLowerCase() ===
            "submitted to billing team" ||
          Boolean(record?.firstTimeBillingRecord?.billingRecordId);
        const resolvedPricingMethod =
          term.pricingMode === "sitewise"
            ? "Site-wise"
            : term.pricingMode === "allocation"
              ? "Allocation Model"
              : "Common";
        const commercialChanged =
          normalizeValue(record?.subscriptionType).toLowerCase() !==
            normalizeValue(mapSubscriptionTypeFromApi(term.subscriptionType)).toLowerCase() ||
          normalizeValue(record?.subscriptionMode).toLowerCase() !==
            normalizeValue(mapSubscriptionModeFromApi(term.subscriptionMode)).toLowerCase() ||
          normalizeValue(record?.subscriptionFee) !==
            normalizeValue(decimalToInputValue(term.subscriptionFee)) ||
          normalizeValue(record?.pricingMethod).toLowerCase() !==
            resolvedPricingMethod.toLowerCase();

        return {
          ...record,
          subscriptionType: mapSubscriptionTypeFromApi(term.subscriptionType),
          subscriptionMode: mapSubscriptionModeFromApi(term.subscriptionMode),
          subscriptionFee: decimalToInputValue(term.subscriptionFee),
          pricingMethod: resolvedPricingMethod,
          ...(completed || !commercialChanged
            ? {}
            : {
                firstTimeBillingValidationStatus: "Billing Calculation",
                firstTimeBillingValidatedAt: "",
                firstTimeBillingValidatedBy: "",
              }),
        };
      });

    setStage2Records((currentRecords) =>
      reconcileCommercialState(currentRecords),
    );
    setStage3Records((currentRecords) =>
      reconcileCommercialState(currentRecords),
    );

    setReadinessRecords((currentRecords) => {
      // Replace only the Screen Codes that were actually saved.
      // Do not remove untouched siblings that share the same Stage 1 group.
      const savedSiteIds = new Set(
        persistedRecords
          .map((record) => normalizeIdInput(record?.screenCode))
          .filter(Boolean),
      );

      const nextRecords = currentRecords.filter(
        (record) =>
          !savedSiteIds.has(normalizeIdInput(record?.screenCode)),
      );

      return [...nextRecords, ...persistedRecords];
    });

    setEditingReadinessRecordId("");
    setEditingReadinessSourceRecordId("");
    setAgreementChoice("");
    setIsAgreementChoiceOpen(false);
    setSiteCreationData(null);
    setMouSubscriptionData(null);

    setSiteCreationResetSignal((value) => value + 1);
    setMouSubscriptionResetSignal((value) => value + 1);
    setIsEntryAreaCollapsed(true);
    onMinimizedChange(true);
    setIsReadinessCollapsed(false);
    shouldRestoreBlankEntryOnExpandRef.current = true;

    window.requestAnimationFrame(() => {
      guidedScrollToElement(readinessSectionRef.current);
    });

    showStage1SaveNotice(
      `${newRecords.length} site record${
        newRecords.length === 1 ? "" : "s"
      } saved to Stage 1 Readiness.`,
    );
  }

  function handleReadinessSelection(recordId, checked) {
    setReadinessRecords((currentRecords) =>
      currentRecords.map((record) => {
        if (record.recordId !== recordId) {
          return record;
        }

        if (!isReadinessRecordSelectable(record)) {
          return { ...record, selected: false };
        }

        if (
          correctionScreenCode &&
          normalizeIdInput(record.screenCode) !== correctionScreenCode
        ) {
          return { ...record, selected: false };
        }

        return { ...record, selected: checked };
      }),
    );
  }

  function handleSelectAllVisibleReadiness() {
    const visibleRecordIds = new Set(
      filteredReadinessRecords
        .filter(
          (record) =>
            isReadinessRecordSelectable(record) &&
            (!correctionScreenCode ||
              normalizeIdInput(record.screenCode) === correctionScreenCode),
        )
        .map((record) => record.recordId),
    );

    setReadinessRecords((currentRecords) =>
      currentRecords.map((record) =>
        visibleRecordIds.has(record.recordId)
          ? { ...record, selected: true }
        : record,
      ),
    );
  }

  function handleToggleVisibleReadinessSelection(checked) {
    const visibleRecordIds = new Set(
      filteredReadinessRecords
        .filter(
          (record) =>
            isReadinessRecordSelectable(record) &&
            (!correctionScreenCode ||
              normalizeIdInput(record.screenCode) === correctionScreenCode),
        )
        .map((record) => record.recordId),
    );

    setReadinessRecords((currentRecords) =>
      currentRecords.map((record) =>
        visibleRecordIds.has(record.recordId)
          ? { ...record, selected: checked }
          : record,
      ),
    );
  }

  function handleClearReadinessSelection() {
    setReadinessRecords((currentRecords) =>
      currentRecords.map((record) => ({ ...record, selected: false })),
    );
  }

  function handleEditReadinessRecord(recordId) {
    const selectedRecord = readinessRecords.find(
      (record) => record.recordId === recordId,
    );

    if (!selectedRecord) {
      return;
    }

    const groupId = getStage1GroupId(selectedRecord);
    const returnedCorrectionScreenCode =
      selectedRecord.stage1CorrectionReturn === true &&
      selectedRecord.complexRenewal !== true
        ? normalizeIdInput(selectedRecord.screenCode)
        : "";

    setCorrectionScreenCode(returnedCorrectionScreenCode);
    setEditingReadinessSourceRecordId(recordId);
    setEditingReadinessRecordId(groupId);
    setIsEntryAreaCollapsed(false);
    onMinimizedChange(false);
    setIsReadinessCollapsed(true);

    window.requestAnimationFrame(() => {
      guidedScrollToElement(mouSubscriptionSectionRef.current);
    });
  }

  function handleCancelStage1Edit() {
    setEditingReadinessRecordId("");
    setEditingReadinessSourceRecordId("");
    setCorrectionScreenCode("");
    setAgreementChoice("");
    setIsAgreementChoiceOpen(false);
    setSiteCreationData(null);
    setMouSubscriptionData(null);
    setSiteCreationResetSignal((value) => value + 1);
    setMouSubscriptionResetSignal((value) => value + 1);
    setIsEntryAreaCollapsed(true);
    onMinimizedChange(true);
    setIsReadinessCollapsed(false);
    shouldRestoreBlankEntryOnExpandRef.current = true;

    window.requestAnimationFrame(() => {
      guidedScrollToElement(readinessSectionRef.current);
    });
  }

  async function executeMoveToStage2(selectedRecords) {
    const existingStage3SiteIds = new Set(
      existingStage3Records.map((record) => normalizeIdInput(record.screenCode)),
    );

    const duplicateStage3Record = selectedRecords.find(
      (record) =>
        !record.stage1CorrectionReturn &&
        existingStage3SiteIds.has(normalizeIdInput(record.screenCode)),
    );

    if (duplicateStage3Record) {
      alert(
        `${duplicateStage3Record.screenCode} already exists in Stage 3. ` +
          "A Screen Code can exist in only one active workflow stage at a time.",
      );
      return;
    }

    const existingStage2SiteIds = new Set(
      existingStage2Records.map((record) => normalizeIdInput(record.screenCode)),
    );
    const duplicateSelectedRecord = selectedRecords.find((record) =>
      existingStage2SiteIds.has(normalizeIdInput(record.screenCode)) &&
      !record.stage1CorrectionReturn,
    );

    if (duplicateSelectedRecord) {
      alert(
        `${duplicateSelectedRecord.screenCode} already exists in Stage 2 Installations.`,
      );
      return;
    }

    const seenSelectedSiteIds = new Set();
    const duplicateWithinSelection = selectedRecords.find((record) => {
      const normalizedSiteId = normalizeIdInput(record.screenCode);
      if (seenSelectedSiteIds.has(normalizedSiteId)) {
        return true;
      }

      seenSelectedSiteIds.add(normalizedSiteId);
      return false;
    });

    if (duplicateWithinSelection) {
      alert(
        `${duplicateWithinSelection.screenCode} is selected more than once for Stage 2.`,
      );
      return;
    }

    const movedRecords = selectedRecords.map((record) => {
      if (record.stage1CorrectionReturn) {
        return {
          ...record,
          manualCompleted: record.manualCompleted ?? false,
          installationCompletionConfirmed:
            record.installationCompletionConfirmed ?? false,
          selected: false,
        };
      }

      return {
        ...record,
        manualCompleted: false,
        installationCompletionConfirmed: false,
        installationCompletedAt: "",
        currentStageStatus: "",
        installationStatus: "",
        dateOfDispatch: "",
        installationDate: "",
        liveDate: "",
        trialPeriod: "",
        trialPeriodExtension: "",
        totalTrialPeriodExtension: "",
        trialExtension: "",
        blockerReason: "",
        remarks: "",
        extensionRemarks: "",
        billingStartDate: "",
        verifiedBy: "",
        selected: false,
        installationDetailsSnapshot: null,
      };
    });

    let persistedSites;

    try {
      persistedSites = await Promise.all(
        selectedRecords
          .filter((record) => record.backendSiteId)
          .map((record) =>
            apiRequest(`/sites/${record.backendSiteId}`, {
              method: "PATCH",
              body: JSON.stringify({
                processingStatus: "Stage 2",
                readinessStatus: "Moved to Stage 2",
                stage1Data: {
                  ...record,
                  processingStatus: "Stage 2",
                  readinessStatus: "Moved to Stage 2",
                  selected: false,
                },
              }),
            }),
          ),
      );
    } catch (error) {
      console.error("Unable to update Stage 1 move status:", error);
      alert(`Unable to move the selected Site to Stage 2. ${error.message}`);
      return;
    }

    /*
     * The PATCH response carries the canonical Site, including any preserved
     * stage2Data. App hydrates the immediate Stage 2 projection from that
     * response instead of waiting for a browser refresh.
     */
    onMoveToStage2(movedRecords, persistedSites);

    if (
      correctionScreenCode &&
      movedRecords.some(
        (record) =>
          normalizeIdInput(record.screenCode) === correctionScreenCode,
      )
    ) {
      setCorrectionScreenCode("");
    }

    setReadinessRecords((currentRecords) =>
      currentRecords.filter((record) => !record.selected),
    );
    setIsReadinessCollapsed(false);

    showStage2MoveNotice(
      `${selectedRecords.length} site record${
        selectedRecords.length === 1 ? "" : "s"
      } moved to Stage 2.`,
    );
  }

  function handleMoveToStage2() {
    const selectedRecords = readinessRecords.filter((record) => {
      if (!record.selected) {
        return false;
      }

      if (!isReadinessRecordSelectable(record)) {
        return false;
      }

      if (
        correctionScreenCode &&
        normalizeIdInput(record.screenCode) !== correctionScreenCode
      ) {
        return false;
      }

      return true;
    });

    if (selectedRecords.length === 0) {
      alert("Select at least one Ready site.");
      return;
    }

    pendingStage2MoveRecordsRef.current = selectedRecords;
    setIsStage2MoveConfirmOpen(true);
  }

  function handleConfirmStage2Move() {
    const selectedRecords = pendingStage2MoveRecordsRef.current;

    setIsStage2MoveConfirmOpen(false);
    pendingStage2MoveRecordsRef.current = [];

    if (selectedRecords.length === 0) {
      return;
    }

    executeMoveToStage2(selectedRecords);
  }

  function handleCancelStage2Move() {
    setIsStage2MoveConfirmOpen(false);
    pendingStage2MoveRecordsRef.current = [];
  }

  return (
    <div className={`site-entry ${isMinimized ? "site-entry--minimized" : ""}`}>
      {stage1SaveNotice ? (
        <div className="site-entry__toast" role="status" aria-live="polite">
          <span className="site-entry__toast-message">{stage1SaveNotice}</span>
          <button
            type="button"
            className="site-entry__toast-dismiss"
            onClick={dismissStage1SaveNotice}
            aria-label="Dismiss save confirmation"
          >
            Dismiss
          </button>
        </div>
      ) : null}

      {stage2MoveNotice ? (
        <div className="site-entry__move-toast" role="status" aria-live="polite">
          <span className="site-entry__move-toast-message">{stage2MoveNotice}</span>
          <button
            type="button"
            className="site-entry__move-toast-dismiss"
            onClick={dismissStage2MoveNotice}
            aria-label="Dismiss Stage 2 move confirmation"
          >
            Dismiss
          </button>
        </div>
      ) : null}

      {isAgreementChoiceOpen ? (
        <div className="site-entry__modal-backdrop" role="presentation">
          <div
            className="site-entry__modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="agreement-choice-title"
          >
            <h3 id="agreement-choice-title">Select Agreement Type</h3>
            <p>
              Agreement type is determined from the existing Billing Code / Customer Code / Complex Code history.
            </p>

            <div className="site-entry__modal-actions">
              <button
                type="button"
                className={
                  agreementChoice === "MoU"
                    ? "site-entry__modal-button site-entry__modal-button--primary"
                    : "site-entry__modal-button"
                }
                onClick={() => setAgreementChoice("MoU")}
              >
                New MoU
              </button>

              <button
                type="button"
                className={
                  agreementChoice === "Addendum"
                    ? "site-entry__modal-button site-entry__modal-button--primary"
                    : "site-entry__modal-button"
                }
                disabled={!addendumAvailable}
                onClick={() => setAgreementChoice("Addendum")}
              >
                Addendum
              </button>
            </div>

            {agreementChoice === "Addendum" ? (
              <div className="site-entry__modal-scope">
                <p><strong>Addendum Scope</strong></p>
                <label>
                  <input
                    type="radio"
                    name="addendumScope"
                    value="Billing Code"
                    checked={addendumScope === "Billing Code"}
                    onChange={() => {
                      setAddendumScope("Billing Code");
                      setAddendumSelectedSiteIds(
                        billingCodeScopeRecords.map((record) => normalizeIdInput(record.screenCode)),
                      );
                    }}
                  />
                  All sites under this Billing Code
                </label>
                <label>
                  <input
                    type="radio"
                    name="addendumScope"
                    value="Selected Sites"
                    checked={addendumScope === "Selected Sites"}
                    onChange={() => {
                      setAddendumScope("Selected Sites");
                      setAddendumSelectedSiteIds(
                        getApplicableSiteRows(siteCreationData?.siteRows || [])
                          .map((row) => normalizeIdInput(row.screenCode))
                          .filter(Boolean),
                      );
                    }}
                  />
                  Selected site(s) only
                </label>

                {addendumScope === "Selected Sites" ? (
                  <div className="site-entry__modal-site-list">
                    {billingCodeScopeRecords.map((record) => {
                      const siteId = normalizeIdInput(record.screenCode);
                      return (
                        <label key={siteId}>
                          <input
                            type="checkbox"
                            checked={addendumSelectedSiteIds.includes(siteId)}
                            onChange={(event) => {
                              setAddendumSelectedSiteIds((current) =>
                                event.target.checked
                                  ? Array.from(new Set([...current, siteId]))
                                  : current.filter((id) => id !== siteId),
                              );
                            }}
                          />
                          {record.screenCode} / {record.screenName} ({getDisplayComplexCode(record.complexCode)})
                        </label>
                      );
                    })}
                  </div>
                ) : null}
              </div>
            ) : null}

            <div className="site-entry__modal-actions">
              <button
                type="button"
                className="site-entry__modal-button site-entry__modal-button--primary"
                onClick={handleConfirmAgreementChoice}
              >
                Continue
              </button>
              <button
                type="button"
                className="site-entry__modal-button"
                onClick={() => {
                  setIsAgreementChoiceOpen(false);
                  setAgreementChoice("");
                }}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      ) : null}

      {isStage2MoveConfirmOpen ? (
        <div className="site-entry__modal-backdrop" role="presentation">
          <div
            className="site-entry__modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="stage2-move-confirm-title"
          >
            <h3 id="stage2-move-confirm-title">Confirm Stage 2 Move</h3>
            <p>Please confirm the entries are correct to proceed further.</p>
            <div className="site-entry__modal-actions">
              <button
                type="button"
                className="site-entry__modal-button site-entry__modal-button--primary"
                onClick={handleConfirmStage2Move}
              >
                Yes
              </button>
              <button
                type="button"
                className="site-entry__modal-button"
                onClick={handleCancelStage2Move}
              >
                No
              </button>
            </div>
          </div>
        </div>
      ) : null}

      <section className="site-entry__header">
        <div />
      </section>

      <section className="site-entry__region">
        <div
          className={`site-entry__region-body ${
            isEntryAreaCollapsed ? "site-entry__region-body--collapsed" : ""
          }`}
        >
          <div className="site-entry__workspace">
            <section className="site-entry__panel" ref={siteCreationSectionRef}>
              <div className="site-entry__panel-header">
                <div>
                  <h3>1A Site Creation</h3>
                </div>
              </div>

              {correctionScreenCode ? (
                <p className="site-entry__panel-note">
                  Correction Screen: <strong>{correctionScreenCode}</strong>. The full
                  complex remains visible for context; sibling screens are locked and
                  will not be resaved or moved forward.
                </p>
              ) : null}

              <SiteCreation
                key={siteCreationResetSignal}
                onDataChange={handleSiteCreationChange}
                resetSignal={siteCreationResetSignal}
                editGroupId={editingReadinessRecordId}
                editGroupRecords={editingReadinessCanonicalRecords}
                blockedSiteIds={blockedStage1SiteIds}
                backendSiteRecords={backendSiteRecords}
                isFullscreenActive={isFullscreenActive}
                onRequestFullscreenRestore={onRequestFullscreenRestore}
                onImportCompleted={handleSiteImportCompleted}
              />
            </section>

            <section
              className="site-entry__panel"
              ref={mouSubscriptionSectionRef}
            >
              <div className="site-entry__panel-header">
                <div>
                  <h3>1B Subscription</h3>
                </div>
              </div>

              <MouSubscription
                key={mouSubscriptionResetSignal}
                siteRows={
                editingReadinessGroupRows.length > 0
                    ? editingReadinessGroupRows
                    : eligibleSiteRowsForStage1
                }
                onDataChange={handleMouSubscriptionChange}
                onSave={handleSaveStage1}
                onRequestAgreementChoice={handleRequestAgreementChoice}
                agreementChoice={agreementChoice}
                agreementChoiceConfirmed={
                  Boolean(agreementChoice) && !isAgreementChoiceOpen
                }
                requiredAgreementType={requiredAgreementType}
                resetSignal={mouSubscriptionResetSignal}
                editGroupId={editingReadinessRecordId}
                editGroupRecords={editingReadinessCanonicalRecords}
                isEditingStage1={Boolean(editingReadinessRecordId)}
                onCancelEdit={handleCancelStage1Edit}
                addendumAvailable={addendumAvailable}
                addendumScope={addendumScope}
                addendumApplicableSiteIds={addendumSelectedSiteIds}
              />
            </section>
          </div>
        </div>
      </section>

      <section
        className="site-entry__readiness"
        ref={readinessSectionRef}
      >
        <div className="site-entry__region-header">
          <div>
            <h3>Stage 1 Readiness</h3>
          {isStage1BackendLoading ? (
            <div className="site-entry__panel-note">Loading saved Stage 1 records...</div>
          ) : null}
          </div>

          <div className="site-entry__region-actions">
            {hasSelectedReadinessRecords ? (
              <button
                type="button"
                className="site-entry__primary-button"
                onClick={handleMoveToStage2}
              >
                Move Selected to Stage 2
              </button>
            ) : null}

            <button
              type="button"
              className="site-entry__secondary-button"
              onClick={() => setIsReadinessCollapsed((current) => !current)}
              aria-expanded={!isReadinessCollapsed}
            >
              {isReadinessCollapsed ? "Expand Readiness" : "Minimize Readiness"}
            </button>
          </div>
        </div>

        <div
          className={`site-entry__region-body ${
            isReadinessCollapsed ? "site-entry__region-body--collapsed" : ""
          }`}
        >
          <div className="site-entry__readiness-controls">
            <label className="site-entry__search">
              <span>Search</span>
              <input
                type="search"
                value={readinessSearchTerm}
                onChange={(event) =>
                  setReadinessSearchTerm(event.target.value)
                }
                placeholder="Billing Code / Customer Code, Complex Code, Screen Code, Screen Name..."
              />
            </label>

            <div className="site-entry__selection-actions">
              <button
                type="button"
                className="site-entry__secondary-button"
                onClick={handleSelectAllVisibleReadiness}
                disabled={filteredReadinessRecords.length === 0}
              >
                Select All
              </button>
              <button
                type="button"
                className="site-entry__secondary-button"
                onClick={handleClearReadinessSelection}
                disabled={readinessRecords.every((record) => !record.selected)}
              >
                Clear Selection
              </button>
            </div>
          </div>

          <div className="site-entry__table-wrapper">
            <table className="site-entry__table">
              <colgroup>
                <col className="site-entry__col-select" />
                <col className="site-entry__col-billing" />
                <col className="site-entry__col-site" />
                <col className="site-entry__col-complex" />
                <col className="site-entry__col-screen" />
                <col className="site-entry__col-location" />
                <col className="site-entry__col-action" />
              </colgroup>
              <thead>
                <tr>
                  <th className="site-entry__table-select-header">
                    <input
                      ref={readinessSelectAllRef}
                      type="checkbox"
                      checked={areAllVisibleReadinessSelected}
                      onChange={(event) =>
                        handleToggleVisibleReadinessSelection(
                          event.target.checked,
                        )
                      }
                      aria-label="Select all visible Stage 1 readiness records"
                      title="Select all"
                      disabled={!hasVisibleReadinessRecords}
                    />
                  </th>
                  <th>Billing Code</th>
                  <th>Screen Code</th>
                  <th>Complex Code</th>
                  <th>Screen Name</th>
                  <th>Location</th>
                  <th className="site-entry__table-action-header">Action</th>
                </tr>
              </thead>

              <tbody>
                {filteredReadinessRecords.length === 0 ? (
                  <tr>
                    <td
                      className="site-entry__table-empty-row"
                      colSpan={7}
                    >
                      No Stage 1 readiness records are available.
                    </td>
                  </tr>
                ) : (
                  filteredReadinessRecords.map((record) => (
                    <tr key={record.recordId}>
                      <td className="site-entry__table-select-cell">
                        <input
                          type="checkbox"
                          checked={record.selected}
                          onChange={(event) =>
                            handleReadinessSelection(
                              record.recordId,
                              event.target.checked,
                            )
                          }
                          aria-label={`Select ${record.screenCode}`}
                          disabled={
                            !isReadinessRecordSelectable(record) ||
                            (Boolean(correctionScreenCode) &&
                              normalizeIdInput(record.screenCode) !==
                                correctionScreenCode)
                          }
                        />
                      </td>

                      <td title={record.billingCode}>{record.billingCode}</td>
                      <td title={record.screenCode}>{record.screenCode}</td>
                      <td title={getDisplayComplexCode(record.complexCode)}>
                        {getDisplayComplexCode(record.complexCode)}
                      </td>
                      <td
                        className="site-entry__screen-cell"
                        title={record.screenName}
                      >
                        {record.screenName}
                      </td>
                      <td
                        className="site-entry__location-cell"
                        title={record.location}
                      >
                        {record.location}
                      </td>
                      <td className="site-entry__table-action-cell">
                        <button
                          type="button"
                          className="site-entry__edit-button"
                          onClick={() => handleEditReadinessRecord(record.recordId)}
                          title="Edit Stage 1 record"
                          disabled={
                            !isReadinessRecordSelectable(record) ||
                            (Boolean(correctionScreenCode) &&
                              normalizeIdInput(record.screenCode) !==
                                correctionScreenCode)
                          }
                        >
                          Edit
                        </button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </section>
    </div>
  );
}

export default SiteEntry;
