import { useEffect, useMemo, useState } from "react";
import { billingApiRequest } from "../../utils/billingApi";
import {
  GOVERNANCE_EVENT_CATEGORIES,
  getGovernanceEventCategory,
} from "../../utils/governanceEventCategory";
import "../GovernanceHistory.css";

const ROWS_PER_PAGE_OPTIONS = [10, 25, 50, 100];

function visiblePageNumbers(currentPage, totalPages) {
  const start = Math.max(1, Math.min(currentPage - 2, totalPages - 4));
  const end = Math.min(totalPages, start + 4);
  return Array.from({ length: end - start + 1 }, (_, index) => start + index);
}

function text(value) {
  return String(value ?? "").trim();
}

function sortedFilterOptions(values) {
  const labels = new Map();
  values.map(text).filter(Boolean).forEach((value) => {
    const key = value.toLocaleLowerCase();
    if (!labels.has(key)) labels.set(key, value);
  });
  return Array.from(labels.values()).sort((left, right) =>
    left.localeCompare(right, undefined, { sensitivity: "base" }),
  );
}

function normalizeSaveMode(event = {}) {
  if (event.saveMode === "Auto" || event.saveMode === "Manual") {
    return event.saveMode;
  }

  const action = text(event.event || event.action);
  const module = text(event.module);
  return module === "ERP" || /sync|recalculat|automatic|expired|renew/i.test(action)
    ? "Auto"
    : "Manual";
}

function eventReason(event = {}) {
  if (text(event.remarks || event.reason)) {
    return text(event.remarks || event.reason);
  }

  const change = event.changes?.[0];
  if (change) {
    return `${change.label || change.field} changed from ${formatHistoryValue(
      change.field,
      change.previous,
    )} to ${formatHistoryValue(change.field, change.next)}.`;
  }

  if (text(event.previousStatus) || text(event.newStatus)) {
    return `${event.event || "Event"}: ${event.previousStatus || "-"} -> ${event.newStatus || "-"}.`;
  }

  return `${event.event || "Business event"} recorded for ${event.screenCode || "the screen"}.`;
}

function parseDate(value) {
  const normalized = text(value);
  if (!normalized) return null;

  const direct = new Date(normalized);
  if (!Number.isNaN(direct.getTime())) {
    return direct;
  }

  return null;
}

function parseJson(value, fallback = {}) {
  if (value && typeof value === "object") {
    return value;
  }

  try {
    return value ? JSON.parse(value) : fallback;
  } catch {
    return fallback;
  }
}

function formatDateTime(value) {
  const date = parseDate(value);

  if (!date) {
    return text(value) || "-";
  }

  return date.toLocaleString("en-IN", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function eventTimestamp(event) {
  return (
    parseDate(event?.eventAt)?.getTime() ||
    parseDate(event?.effectiveDate)?.getTime() ||
    0
  );
}

function addEvent(events, event) {
  if (!event?.screenCode || !event?.event) {
    return;
  }

  events.push({
    eventId:
      event.eventId ||
      [
        event.screenCode,
        event.module,
        event.event,
        event.eventAt || event.effectiveDate || "",
      ].join("::"),
    billingCode: text(event.billingCode),
    complexCode: text(event.complexCode),
    screenCode: text(event.screenCode),
    screenName: text(event.screenName),
    location: text(event.location),
    module: text(event.module) || "Workflow",
    event: text(event.event),
    previousStatus: text(event.previousStatus),
    newStatus: text(event.newStatus),
    remarks: eventReason(event),
    changedBy: text(event.changedBy),
    workflowType: text(event.workflowType),
    changes: Array.isArray(event.changes) ? event.changes : [],
    eventAt: text(event.eventAt || event.effectiveDate),
    saveMode: normalizeSaveMode(event),
  });
}

const historyFieldLabels = {
  pricingMode: "Pricing Mode",
  subscriptionType: "Subscription Type",
  subscriptionMode: "Subscription Mode",
  subscriptionFee: "Subscription Fee",
  otfApplicable: "OTF Applicable",
  otfType: "OTF Type",
  otfAmount: "OTF Amount",
  otfTaxMode: "OTF Tax Mode",
  processingStatus: "Processing Status",
  readinessStatus: "Readiness Status",
};

function formatHistoryValue(field, value) {
  if (value === null || value === undefined || value === "") return "-";
  if (["subscriptionFee", "otfAmount"].includes(field)) {
    const number = Number(String(value).replace(/,/g, ""));
    if (Number.isFinite(number)) {
      return new Intl.NumberFormat("en-IN", {
        style: "currency",
        currency: "INR",
        maximumFractionDigits: 2,
      }).format(number);
    }
  }
  return text(value);
}

function compactEventName(row) {
  if (row.module === "Commercial" && row.event === "Commercial Terms Updated") {
    return "Terms Updated";
  }
  return row.event;
}

function transitionText(row) {
  const firstChange = row.changes?.[0];
  if (firstChange) {
    return `${formatHistoryValue(firstChange.field, firstChange.previous)} -> ${formatHistoryValue(firstChange.field, firstChange.next)}`;
  }
  return `${row.previousStatus || "-"} -> ${row.newStatus || "-"}`;
}

function addCommonLifecycleEvents(events, record, moduleName) {
  const base = {
    billingCode: record?.billingCode,
    complexCode: record?.complexCode,
    screenCode: record?.screenCode,
    screenName: record?.screenName,
    location: record?.location,
    module: moduleName,
  };

  if (record?.savedAt) {
    addEvent(events, {
      ...base,
      event: "Record Saved",
      newStatus: record?.processingStatus || record?.currentStageStatus || "",
      remarks: record?.remarks || "",
      eventAt: record.savedAt,
    });
  }

  if (record?.installationCompletedAt) {
    addEvent(events, {
      ...base,
      module: "Installations",
      event: "Installation Details Saved",
      newStatus:
        record?.installationStatus || record?.currentStageStatus || "Installed",
      remarks: record?.remarks || "",
      eventAt: record.installationCompletedAt,
      changedBy: record?.verifiedBy || "",
    });
  }

  if (record?.billingVerifiedAt) {
    addEvent(events, {
      ...base,
      module: "Billings",
      event: "Billing Verified",
      previousStatus: "First Billing Pending Approval",
      newStatus:
        record?.billingVerificationStatus || "Billing Verified",
      remarks: record?.billingRemarks || "",
      eventAt: record.billingVerifiedAt,
    });
  }

  if (record?.commercialUpdatedAt) {
    addEvent(events, {
      ...base,
      module: "Commercial",
      event: "Commercial Terms Updated",
      newStatus:
        record?.commercialStatus ||
        record?.pricingMethod ||
        record?.billingMode ||
        "",
      remarks: record?.addendumRemarks || "",
      eventAt: record.commercialUpdatedAt,
    });
  }

  const closureHistory = Array.isArray(record?.closureHistory)
    ? record.closureHistory
    : [];

  closureHistory.forEach((entry, index) => {
    const newStatus = text(entry?.newStatus);
    const isActive = newStatus.toLowerCase() === "active";

    addEvent(events, {
      ...base,
      module: "Billings",
      eventId:
        entry?.eventId ||
        `${record?.screenCode || "screen"}::closure::${index}`,
      event: isActive ? "Site Reactivated" : "Billing Status Changed",
      previousStatus: entry?.previousStatus || "",
      newStatus,
      remarks:
        entry?.remarks ||
        entry?.reason ||
        entry?.restoreRemarks ||
        "",
      changedBy: entry?.updatedBy || entry?.changedBy || "",
      eventAt:
        entry?.updatedAt ||
        entry?.changedAt ||
        entry?.effectiveDate ||
        "",
    });
  });

  const commercialHistory = Array.isArray(record?.commercialHistory)
    ? record.commercialHistory
    : [];

  commercialHistory.forEach((entry, index) => {
    addEvent(events, {
      ...base,
      module: "Commercial",
      eventId:
        entry?.eventId ||
        `${record?.screenCode || "screen"}::commercial::${index}`,
      event: entry?.event || "Commercial Status Changed",
      previousStatus:
        entry?.previousStatus ||
        entry?.previousCommercialStatus ||
        "",
      newStatus:
        entry?.newStatus ||
        entry?.commercialStatus ||
        "",
      remarks: entry?.remarks || "",
      changedBy: entry?.changedBy || entry?.updatedBy || "",
      eventAt:
        entry?.changedAt ||
        entry?.updatedAt ||
        entry?.effectiveDate ||
        "",
    });
  });
}

function buildHistory({
  stage1Records,
  stage2Records,
  stage3Records,
  billingRecords,
  incentiveStates,
  governanceAuditRows = [],
}) {
  const events = [];

  (stage1Records || []).forEach((record) => {
    addCommonLifecycleEvents(events, record, "Site Entry");

    if (record?.requestedAt) {
      addEvent(events, {
        billingCode: record.billingCode,
        complexCode: record.complexCode,
        screenCode: record.screenCode,
        screenName: record.screenName,
        location: record.location,
        module: "Site Entry",
        event: record?.correctionType
          ? `${record.correctionType} Requested`
          : "Stage 1 Correction Requested",
        newStatus: record?.correctionRoute || "",
        remarks: record?.addendumRemarks || "",
        eventAt: record.requestedAt,
      });
    }
  });

  (stage2Records || []).forEach((record) => {
    addCommonLifecycleEvents(events, record, "Installations");

    if (record?.installationDate) {
      addEvent(events, {
        billingCode: record.billingCode,
        complexCode: record.complexCode,
        screenCode: record.screenCode,
        screenName: record.screenName,
        location: record.location,
        module: "Installations",
        event: "Installation Date Recorded",
        newStatus: record?.installationStatus || "",
        remarks: record?.remarks || "",
        eventAt: record.installationDate,
      });
    }

    if (record?.billingStartDate) {
      addEvent(events, {
        billingCode: record.billingCode,
        complexCode: record.complexCode,
        screenCode: record.screenCode,
        screenName: record.screenName,
        location: record.location,
        module: "Installations",
        event: "Billing Start Date Confirmed",
        newStatus: "Billing Scheduled",
        remarks: record?.remarks || "",
        eventAt: record.billingStartDate,
      });
    }
  });

  (stage3Records || []).forEach((record) => {
    addCommonLifecycleEvents(events, record, "Billings");

    if (record?.stage3BillingStartDateSnapshot) {
      addEvent(events, {
        billingCode: record.billingCode,
        complexCode: record.complexCode,
        screenCode: record.screenCode,
        screenName: record.screenName,
        location: record.location,
        module: "Workflow",
        event: "Moved to Billings",
        previousStatus: "Installations",
        newStatus: "Billings",
        eventAt:
          record?.billingVerifiedAt ||
          record?.lastUpdatedAt ||
          record?.stage3BillingStartDateSnapshot,
      });
    }
  });

  (billingRecords || []).forEach((record) => {
    const entries = Array.isArray(record?.invoiceEntries)
      ? record.invoiceEntries
      : [];

    entries.forEach((entry, index) => {
      if (entry?.invoiceNumber || entry?.invoiceDate) {
        addEvent(events, {
          billingCode: record.billingCode,
          complexCode: record.complexCode,
          screenCode: record.screenCode || record.siteScope,
          screenName: record.screenName,
          location: record.location,
          module: "Billings",
          eventId:
            entry?.entryId ||
            `${record?.billingRecordId || "billing"}::invoice::${index}`,
          event: "Invoice Updated",
          newStatus: entry?.paymentStatus || record?.paymentStatus || "",
          remarks: entry?.remarks || record?.billingRemarks || "",
          eventAt:
            entry?.invoiceDate ||
            entry?.updatedAt ||
            record?.invoiceDate ||
            "",
        });
      }

      if (
        String(entry?.paymentStatus || "").toLowerCase() === "paid" ||
        entry?.receiptNumber ||
        entry?.paymentReceivedDate
      ) {
        addEvent(events, {
          billingCode: record.billingCode,
          complexCode: record.complexCode,
          screenCode: record.screenCode || record.siteScope,
          screenName: record.screenName,
          location: record.location,
          module: "Billings",
          eventId:
            `${entry?.entryId || index}::payment`,
          event: "Payment Received",
          previousStatus: "Pending",
          newStatus: "Paid",
          remarks: entry?.paymentRemarks || entry?.remarks || "",
          eventAt:
            entry?.paymentReceivedDate ||
            entry?.receiptDate ||
            entry?.updatedAt ||
            "",
        });
      }
    });
  });

  Object.entries(incentiveStates || {}).forEach(
    ([screenCode, state]) => {
      const downloadedParts =
        state?.downloadedParts && typeof state.downloadedParts === "object"
          ? state.downloadedParts
          : {};

      Object.entries(downloadedParts).forEach(([partKey, part]) => {
        const stageRecord = (stage3Records || []).find(
          (record) => text(record?.screenCode) === text(screenCode),
        );

        addEvent(events, {
          billingCode: stageRecord?.billingCode,
          complexCode: stageRecord?.complexCode,
          screenCode,
          screenName: stageRecord?.screenName,
          location: stageRecord?.location,
          module: "Incentive",
          event: "Incentive Downloaded",
          newStatus: partKey,
          remarks: "",
          eventAt: part?.downloadedAt || "",
        });
      });
    },
  );

  const safeGovernanceAuditRows = Array.isArray(governanceAuditRows)
    ? governanceAuditRows.filter(
        (audit) => audit && typeof audit === "object" && !Array.isArray(audit),
      )
    : [];

  safeGovernanceAuditRows.forEach((audit) => {
    const previous = parseJson(audit.previousValue, {});
    const next = parseJson(audit.newValue, {});
    const isCorrection =
      text(audit?.module) === "Governance" &&
      text(audit?.action) === "Record Correction";
    const isCommercialUpdate =
      text(audit?.action) === "Commercial Terms Updated";
    const isDedicatedStage3Movement =
      text(audit?.module) === "Workflow" &&
      text(audit?.action) === "Moved to Stage 3";
    const isSnapshotStage3Movement =
      text(audit?.module) === "Site Entry" &&
      text(audit?.action) === "Stage 1 Site Updated" &&
      (text(previous.processingStatus).toLowerCase() === "stage 2" ||
        text(previous.readinessStatus).toLowerCase() === "moved to stage 2") &&
      (text(next.processingStatus).toLowerCase() === "stage 3" ||
        text(next.readinessStatus).toLowerCase() === "moved to stage 3");
    const isMovedToStage3 =
      isDedicatedStage3Movement || isSnapshotStage3Movement;
    const isDedicatedStage2Return =
      text(audit?.module) === "Workflow" &&
      text(audit?.action) === "Returned to Stage 2";
    const isSnapshotStage2Return =
      text(audit?.module) === "Site Entry" &&
      text(audit?.action) === "Stage 1 Site Updated" &&
      (text(previous.processingStatus).toLowerCase() === "stage 3" ||
        text(previous.readinessStatus).toLowerCase() === "moved to stage 3") &&
      (text(next.processingStatus).toLowerCase() === "stage 2" ||
        text(next.readinessStatus).toLowerCase() === "moved to stage 2");
    const isReturnedToStage2 =
      isDedicatedStage2Return || isSnapshotStage2Return;
    const isStageMovement =
      isMovedToStage3 || isReturnedToStage2;

    const parsedFields = parseJson(audit.field, []);
    const changedFields = Array.isArray(parsedFields)
      ? parsedFields
      : Array.isArray(previous?.changedFields)
        ? previous.changedFields
        : [];

    if (isStageMovement) {
      addEvent(events, {
        eventId: audit.id,
        billingCode: audit.site?.billingId,
        complexCode: audit.site?.complexId,
        screenCode: audit.site?.siteId,
        screenName: audit.site?.screenName,
        location: audit.site?.location,
        module: "Workflow",
        event: isReturnedToStage2
          ? "Returned to Stage 2"
          : "Moved to Stage 3",
        previousStatus: isReturnedToStage2 ? "Stage 3" : "Stage 2",
        newStatus: isReturnedToStage2 ? "Stage 2" : "Stage 3",
        changes: [
          {
            field: "stage",
            label: "Stage",
            previous: isReturnedToStage2 ? "Stage 3" : "Stage 2",
            next: isReturnedToStage2 ? "Stage 2" : "Stage 3",
          },
        ],
        remarks: audit.remarks || audit.reason,
        changedBy: audit.changedBy,
        eventAt: audit.createdAt,
        saveMode: audit.saveMode,
      });
      return;
    }

    if (!isCorrection && !isCommercialUpdate) {
      addEvent(events, {
        eventId: audit.id,
        billingCode: audit.site?.billingId,
        complexCode: audit.site?.complexId,
        screenCode: audit.site?.siteId,
        screenName: audit.site?.screenName,
        location: audit.site?.location,
        module: audit.module,
        event: audit.action,
        previousStatus: audit.previousValue,
        newStatus: audit.newValue,
        remarks: audit.remarks || audit.reason,
        changedBy: audit.changedBy,
        eventAt: audit.createdAt,
        saveMode: audit.saveMode,
      });
      return;
    }

    if (isCommercialUpdate) {
      const changes = changedFields
        .filter(
          (field) =>
            Object.prototype.hasOwnProperty.call(previous, field) ||
            Object.prototype.hasOwnProperty.call(next, field),
        )
        .map((field) => ({
          field,
          label: historyFieldLabels[field] || field,
          previous: previous[field],
          next: next[field],
        }));
      const firstChange = changes[0];

      addEvent(events, {
        eventId: audit.id,
        billingCode: audit.site?.billingId,
        complexCode: audit.site?.complexId,
        screenCode: audit.site?.siteId,
        screenName: audit.site?.screenName,
        location: audit.site?.location,
        module: "Commercial",
        event: "Commercial Terms Updated",
        previousStatus: firstChange?.previous,
        newStatus: firstChange?.next,
        changes,
        remarks: audit.remarks || audit.reason,
        changedBy: audit.changedBy,
        eventAt: audit.createdAt,
        saveMode: audit.saveMode,
      });
      return;
    }

    const previousStage = previous?.previousValues?.processingStatus || "";

    if (text(previousStage).toLowerCase() !== "stage 3") {
      return;
    }

    addEvent(events, {
      eventId: previous?.eventId || audit.id,
      billingCode: audit.site?.billingId,
      complexCode: audit.site?.complexId,
      screenCode: audit.site?.siteId,
      screenName: audit.site?.screenName,
      module: "Record Correction",
      event: "Returned to Stage 1",
      previousStatus: "Stage 3",
      newStatus: "Stage 1",
      remarks: audit.remarks || audit.reason,
      changedBy: audit.changedBy,
      workflowType: previous?.workflowType || "RECORD_CORRECTION",
      changes: [
        {
          field: "stage",
          label: "Stage",
          previous: "Stage 3",
          next: "Stage 1",
        },
      ],
      eventAt: audit.createdAt,
      saveMode: audit.saveMode,
    });
  });

  const persistedCommercialAuditIds = new Set(
    safeGovernanceAuditRows
      .filter((audit) => text(audit?.action) === "Commercial Terms Updated")
      .map((audit) => text(audit?.id))
      .filter(Boolean),
  );
  const persistedCommercialScreens = new Set(
    safeGovernanceAuditRows
      .filter((audit) => text(audit?.action) === "Commercial Terms Updated")
      .map((audit) => text(audit?.site?.siteId))
      .filter(Boolean),
  );
  const displayEvents = events.filter(
    (event) =>
      event.module !== "Commercial" ||
      event.event !== "Commercial Terms Updated" ||
      !persistedCommercialScreens.has(event.screenCode) ||
      persistedCommercialAuditIds.has(event.eventId),
  );

  const uniqueEvents = Array.from(
    displayEvents.reduce((map, event) => {
      const key =
        event.eventId ||
        [
          event.screenCode,
          event.module,
          event.event,
          event.eventAt,
          event.newStatus,
        ].join("::");

      if (!map.has(key)) {
        map.set(key, event);
      }

      return map;
    }, new Map()).values(),
  );

  return uniqueEvents.sort(
    (left, right) => eventTimestamp(right) - eventTimestamp(left),
  );
}

export default function HistoryTracking({
  stage1Records = [],
  stage2Records = [],
  stage3Records = [],
  billingRecords = [],
  incentiveStates = {},
}) {
  const [search, setSearch] = useState("");
  const [moduleFilter, setModuleFilter] = useState("All");
  const [eventFilter, setEventFilter] = useState("All");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [activeEvent, setActiveEvent] = useState(null);
  const [rowsPerPage, setRowsPerPage] = useState(10);
  const [currentPage, setCurrentPage] = useState(1);
  const [activeScreenCode, setActiveScreenCode] = useState("");
  const [governanceAuditRows, setGovernanceAuditRows] = useState([]);
  const [governanceAuditError, setGovernanceAuditError] = useState("");

  useEffect(() => {
    let cancelled = false;

    billingApiRequest("/audit-history")
      .then((auditRows) => {
        if (!cancelled) {
          setGovernanceAuditRows(Array.isArray(auditRows) ? auditRows : []);
          setGovernanceAuditError("");
        }
      })
      .catch((error) => {
        if (!cancelled) {
          setGovernanceAuditError(
            `Persisted Governance history could not be loaded. ${
              error?.message || "Please try again later."
            }`,
          );
          console.error("Unable to load persisted Governance history:", error);
        }
      });

    return () => {
      cancelled = true;
    };
  }, []);

  const rows = useMemo(
    () =>
      buildHistory({
        stage1Records,
        stage2Records,
        stage3Records,
        billingRecords,
        incentiveStates,
        governanceAuditRows,
      }),
    [
      stage1Records,
      stage2Records,
      stage3Records,
      billingRecords,
      incentiveStates,
      governanceAuditRows,
    ],
  );

  const filteredRows = useMemo(() => {
    const query = text(search).toLowerCase();
    const fromDate = dateFrom ? new Date(`${dateFrom}T00:00:00`) : null;
    const toDate = dateTo ? new Date(`${dateTo}T23:59:59`) : null;

    return rows.filter((row) => {
      if (moduleFilter !== "All" && text(row.module) !== moduleFilter) {
        return false;
      }
      if (
        eventFilter !== "All" &&
        getGovernanceEventCategory(row.event) !== eventFilter
      ) {
        return false;
      }

      const rowDate = parseDate(row.eventAt);

      if (fromDate && rowDate && rowDate < fromDate) {
        return false;
      }

      if (toDate && rowDate && rowDate > toDate) {
        return false;
      }

      if (!query) {
        return true;
      }

      return [
        row.billingCode,
        row.complexCode,
        row.screenCode,
        row.screenName,
        row.location,
        row.module,
        row.event,
        row.previousStatus,
        row.newStatus,
        row.remarks,
      ]
        .join(" ")
        .toLowerCase()
        .includes(query);
    });
  }, [rows, search, moduleFilter, eventFilter, dateFrom, dateTo]);

  const moduleOptions = useMemo(
    () => sortedFilterOptions(rows.map((row) => row.module)),
    [rows],
  );
  const eventOptions = GOVERNANCE_EVENT_CATEGORIES;

  useEffect(() => {
    setCurrentPage(1);
  }, [search, moduleFilter, eventFilter, dateFrom, dateTo, rowsPerPage]);

  const totalPages = Math.max(1, Math.ceil(filteredRows.length / rowsPerPage));

  useEffect(() => {
    setCurrentPage((page) => Math.min(page, totalPages));
  }, [totalPages]);

  const pagedRows = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filteredRows.slice(start, start + rowsPerPage);
  }, [currentPage, filteredRows, rowsPerPage]);

  const firstVisible = filteredRows.length
    ? (currentPage - 1) * rowsPerPage + 1
    : 0;
  const lastVisible = Math.min(currentPage * rowsPerPage, filteredRows.length);

  const screenHistory = activeScreenCode
    ? rows.filter(
        (row) => text(row.screenCode) === text(activeScreenCode),
      )
    : [];

  return (
    <div className="governance-history">
      <div style={{ marginBottom: 18 }}>
        <div className="governance-history__eyebrow">
          Governance
        </div>
        <h1>History Tracking</h1>
        <p className="governance-history__intro">
          Screen-wise workflow and lifecycle history. Most recent event is
          shown first.
        </p>
        {governanceAuditError ? (
          <p
            role="status"
            style={{
              margin: "10px 0 0",
              padding: "8px 10px",
              border: "1px solid #d9dee7",
              borderRadius: 8,
              color: "#4b5563",
              background: "#f7f7f8",
            }}
          >
            {governanceAuditError}
          </p>
        ) : null}
      </div>

      <div className="governance-history__filters">
        <input
          className="governance-history__control"
          type="search"
          placeholder="Search Billing ID, Complex ID, Screen, Location, Event..."
          value={search}
          onChange={(event) => setSearch(event.target.value)}
        />

        <select
          className="governance-history__control"
          value={moduleFilter}
          onChange={(event) => setModuleFilter(event.target.value)}
        >
          <option value="All">All Modules</option>
          {moduleOptions.map((moduleName) => (
            <option key={moduleName} value={moduleName}>{moduleName}</option>
          ))}
        </select>

        <select
          className="governance-history__control"
          value={eventFilter}
          onChange={(event) => setEventFilter(event.target.value)}
          aria-label="Filter by event"
        >
          <option value="All">All Events</option>
          {eventOptions.map((category) => (
            <option key={category} value={category}>{category}</option>
          ))}
        </select>

        <input
          className="governance-history__control"
          type="date"
          placeholder="dd-mm-yyyy"
          value={dateFrom}
          onChange={(event) => setDateFrom(event.target.value)}
          title="From date"
        />

        <input
          className="governance-history__control"
          type="date"
          placeholder="dd-mm-yyyy"
          value={dateTo}
          onChange={(event) => setDateTo(event.target.value)}
          title="To date"
        />

        <label className="governance-history__rpp-filter">
          RPP
          <select
            className="governance-history__control governance-history__rpp"
            aria-label="Rows per page"
            value={rowsPerPage}
            onChange={(event) => {
              setRowsPerPage(Number(event.target.value));
              setCurrentPage(1);
            }}
          >
            {ROWS_PER_PAGE_OPTIONS.map((option) => (
              <option key={option} value={option}>{option}</option>
            ))}
          </select>
        </label>

        <button
          type="button"
          className="governance-history__button"
          onClick={() => {
            setSearch("");
            setModuleFilter("All");
            setEventFilter("All");
            setDateFrom("");
            setDateTo("");
            setCurrentPage(1);
          }}
        >
          Reset
        </button>
      </div>

      <div className="governance-history__card">
        <div className="governance-history__card-header">
          <strong>History Records</strong>
          <span className="governance-history__count">
            {filteredRows.length} event{filteredRows.length === 1 ? "" : "s"}
          </span>
        </div>

        <div className="governance-history__table-wrap">
          <table className="governance-history__table">
            <colgroup>
              <col />
              <col />
              <col />
              <col />
              <col />
            </colgroup>
            <thead>
              <tr>
                {[
                  "Date / Time",
                  "Screen Code",
                  "Module",
                  "Event",
                  "View",
                ].map((heading) => (
                  <th key={heading}>{heading}</th>
                ))}
              </tr>
            </thead>

            <tbody>
              {filteredRows.length ? (
                pagedRows.map((row) => (
                  <tr key={row.eventId}>
                    <td>{formatDateTime(row.eventAt)}</td>
                    <td><span className="governance-history__ellipsis" title={row.screenCode}>{row.screenCode || "-"}</span></td>
                    <td><span className="governance-history__ellipsis" title={row.module}>{row.module}</span></td>
                    <td><span className="governance-history__ellipsis" title={compactEventName(row)}>{compactEventName(row)}</span></td>
                    <td>
                      <button
                        type="button"
                        onClick={() => setActiveEvent(row)}
                        className="governance-history__button governance-history__view"
                      >
                        View
                      </button>
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={5} className="governance-history__empty">
                    No history records found.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        {filteredRows.length ? (
          <div className="governance-history__pagination">
            <div className="governance-history__pagination-summary">
              <span>Showing {firstVisible}-{lastVisible} of {filteredRows.length}</span>
            </div>
            <div className="governance-history__pagination-actions">
              <button type="button" className="governance-history__button" disabled={currentPage <= 1} onClick={() => setCurrentPage((page) => Math.max(1, page - 1))}>Previous</button>
              {visiblePageNumbers(currentPage, totalPages).map((page) => (
                <button
                  type="button"
                  key={page}
                  className={`governance-history__page-button ${page === currentPage ? "governance-history__page-button--active" : ""}`}
                  aria-current={page === currentPage ? "page" : undefined}
                  onClick={() => setCurrentPage(page)}
                >
                  {page}
                </button>
              ))}
              <button type="button" className="governance-history__button" disabled={currentPage >= totalPages} onClick={() => setCurrentPage((page) => Math.min(totalPages, page + 1))}>Next</button>
            </div>
          </div>
        ) : null}
      </div>

      {activeEvent ? (
        <div className="governance-history__backdrop" role="presentation" onMouseDown={() => setActiveEvent(null)}>
          <section
            className="governance-history__dialog"
            role="dialog"
            aria-modal="true"
            aria-labelledby="history-event-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="governance-history__dialog-header">
              <div>
                <div className="governance-history__screen">{activeEvent.screenCode}</div>
                <h2 id="history-event-title">{activeEvent.event}</h2>
              </div>
              <button type="button" className="governance-history__button" onClick={() => setActiveEvent(null)}>Close</button>
            </div>

            {activeEvent.changes?.length ? (
              <div className="governance-history__changes">
                {activeEvent.changes.map((change, index) => (
                <div className="governance-history__change" key={`${change.field || change.label}-${index}`}>
                  <div className="governance-history__change-label">{change.label || historyFieldLabels[change.field] || change.field}</div>
                  <div className="governance-history__transition">
                    {formatHistoryValue(change.field, change.previous)} -&gt; {formatHistoryValue(change.field, change.next)}
                  </div>
                </div>
                ))}
              </div>
            ) : (
              <div className="governance-history__detail-empty">No field-level changes recorded</div>
            )}

            <div className="governance-history__metadata">
              <div><span>Screen Code</span><br />{activeEvent.screenCode || "-"}</div>
              <div><span>Module</span><br />{activeEvent.module || "-"}</div>
              <div><span>Event / Action</span><br />{activeEvent.event || "-"}</div>
              <div><span>Billing Code</span><br />{activeEvent.billingCode || "-"}</div>
              <div><span>Complex Code</span><br />{activeEvent.complexCode || "Standalone"}</div>
              <div><span>Screen Name</span><br />{activeEvent.screenName || "-"}</div>
              <div><span>Date &amp; Time</span><br />{formatDateTime(activeEvent.eventAt)}</div>
              <div><span>Save Mode</span><br />{activeEvent.saveMode || "Not Available"}</div>
              <div><span>Reason</span><br />{activeEvent.remarks || "-"}</div>
              <div><span>From -&gt; To</span><br />{transitionText(activeEvent)}</div>
              <div><span>Changed By</span><br />{activeEvent.changedBy || "Not Available"}</div>
              <div><span>Location</span><br />{activeEvent.location || "-"}</div>
            </div>

            <div className="governance-history__dialog-actions">
              <button
                type="button"
                className="governance-history__button"
                onClick={() => {
                  setActiveScreenCode(activeEvent.screenCode);
                  setActiveEvent(null);
                }}
              >
                Screen History
              </button>
            </div>
          </section>
        </div>
      ) : null}

      {activeScreenCode ? (
        <div
          style={{
            marginTop: 18,
            border: "1px solid #e1e5ec",
            borderRadius: 12,
            background: "#ffffff",
            padding: 16,
          }}
        >
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              gap: 12,
              alignItems: "center",
              marginBottom: 12,
            }}
          >
            <div>
              <h2 style={{ margin: 0, fontSize: 20 }}>
                Screen History — {activeScreenCode}
              </h2>
              <div style={{ color: "#667085", marginTop: 4 }}>
                Most recent event first
              </div>
            </div>

            <button
              type="button"
              onClick={() => setActiveScreenCode("")}
              style={{
                border: "1px solid #cfd5df",
                borderRadius: 7,
                padding: "7px 12px",
                background: "#ffffff",
                cursor: "pointer",
              }}
            >
              Close
            </button>
          </div>

          <div style={{ display: "grid", gap: 8 }}>
            {screenHistory.map((row) => (
              <div
                key={`detail-${row.eventId}`}
                style={{
                  border: "1px solid #edf0f4",
                  borderRadius: 9,
                  padding: 12,
                  display: "grid",
                  gridTemplateColumns: "170px 150px 1fr",
                  gap: 12,
                }}
              >
                <div>
                  <strong>{formatDateTime(row.eventAt)}</strong>
                </div>
                <div>{row.module}</div>
                <div>
                  <strong>{row.event}</strong>
                  <div style={{ marginTop: 4, color: "#667085" }}>
                    {row.previousStatus || "-"} → {row.newStatus || "-"}
                  </div>
                  {row.remarks ? (
                    <div style={{ marginTop: 4 }}>{row.remarks}</div>
                  ) : null}
                </div>
              </div>
            ))}
          </div>
        </div>
      ) : null}
    </div>
  );
}
