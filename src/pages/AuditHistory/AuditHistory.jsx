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

function n(value) {
  return String(value ?? "").trim();
}

function sortedFilterOptions(values) {
  const labels = new Map();
  values.map(n).filter(Boolean).forEach((value) => {
    const key = value.toLocaleLowerCase();
    if (!labels.has(key)) labels.set(key, value);
  });
  return Array.from(labels.values()).sort((left, right) =>
    left.localeCompare(right, undefined, { sensitivity: "base" }),
  );
}

function normalizeSaveMode(entry = {}) {
  if (entry.saveMode === "Auto" || entry.saveMode === "Manual") {
    return entry.saveMode;
  }

  const action = n(entry.action);
  const module = n(entry.module);
  return module === "ERP" || /sync|recalculat|automatic|expired|renew/i.test(action)
    ? "Auto"
    : "Manual";
}

function parseDate(value) {
  const text = n(value);
  if (!text) return null;
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? null : date;
}

const auditFieldLabels = {
  pricingMode: "Pricing Mode",
  subscriptionType: "Subscription Type",
  subscriptionMode: "Subscription Mode",
  subscriptionFee: "Subscription Fee",
  otfApplicable: "OTF Applicable",
  otfType: "OTF Type",
  otfAmount: "OTF Amount",
  otfTaxMode: "OTF Tax Mode",
};

function formatAuditValue(value, changedFields = []) {
  const parsed = parseJson(value, null);
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    return n(value);
  }

  const fields =
    Array.isArray(changedFields) && changedFields.length > 0
      ? changedFields
      : Object.keys(parsed);

  return fields
    .filter((field) => Object.prototype.hasOwnProperty.call(parsed, field))
    .map((field) => `${auditFieldLabels[field] || field}: ${parsed[field] ?? "-"}`)
    .join(" | ");
}

function formatDateTime(value) {
  const date = parseDate(value);
  if (!date) return n(value) || "-";

  return date.toLocaleString("en-IN", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function compactAuditAction(action) {
  if (action === "Commercial Terms Updated") return "Terms Updated";
  if (action === "Record Correction") return "Returned to Stage 1";
  return action;
}

function formatDetailValue(field, value) {
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
  return n(value);
}

function detailChanges(row) {
  const fields = Array.isArray(row?.changedFields) && row.changedFields.length
    ? row.changedFields
    : [];
  const previous = parseJson(row?.rawPreviousValue, null);
  const next = parseJson(row?.rawNewValue, null);

  if (
    fields.length &&
    previous && typeof previous === "object" && !Array.isArray(previous) &&
    next && typeof next === "object" && !Array.isArray(next)
  ) {
    return fields.map((field) => ({
      field,
      label: auditFieldLabels[field] || field,
      previous: previous[field],
      next: next[field],
    }));
  }

  if (
    !fields.length &&
    (n(row?.module) === "Workflow" || /stage|return|move/i.test(n(row?.action)))
  ) {
    return [];
  }

  return [{
    field: n(row?.field) || "value",
    label: n(row?.field) || "Value",
    previous: row?.previousValue,
    next: row?.newValue,
  }];
}

function addAudit(rows, entry) {
  if (!n(entry?.screenCode) || !n(entry?.action)) return;

  rows.push({
    auditId:
      n(entry?.auditId) ||
      [
        n(entry?.screenCode),
        n(entry?.module),
        n(entry?.action),
        n(entry?.changedAt),
        rows.length,
      ].join("::"),
    changedAt: n(entry?.changedAt),
    billingCode: n(entry?.billingCode),
    complexCode: n(entry?.complexCode),
    screenCode: n(entry?.screenCode),
    screenName: n(entry?.screenName),
    module: n(entry?.module) || "Workflow",
    action: n(entry?.action),
    workflowType: n(entry?.workflowType),
    field: n(entry?.field),
    previousValue: formatAuditValue(entry?.previousValue, entry?.changedFields),
    newValue: formatAuditValue(entry?.newValue, entry?.changedFields),
    rawPreviousValue: entry?.rawPreviousValue ?? entry?.previousValue,
    rawNewValue: entry?.rawNewValue ?? entry?.newValue,
    changedFields: Array.isArray(entry?.changedFields) ? entry.changedFields : [],
    changedBy: n(entry?.changedBy) || "Not Available",
    remarks: n(entry?.remarks || entry?.reason),
    saveMode: normalizeSaveMode(entry),
  });
}

function parseJson(value, fallback = {}) {
  try {
    return value ? JSON.parse(value) : fallback;
  } catch {
    return fallback;
  }
}

function collectAuditRows({
  stage1Records,
  stage2Records,
  stage3Records,
  billingRecords,
  incentiveStates,
  governanceAuditRows = [],
}) {
  const rows = [];

  const allStageRecords = [
    ...(stage1Records || []),
    ...(stage2Records || []),
    ...(stage3Records || []),
  ];

  allStageRecords.forEach((record) => {
    const base = {
      billingCode: record?.billingCode,
      complexCode: record?.complexCode,
      screenCode: record?.screenCode,
      screenName: record?.screenName,
    };

    const closureHistory = Array.isArray(record?.closureHistory)
      ? record.closureHistory
      : [];

    closureHistory.forEach((entry, index) => {
      addAudit(rows, {
        ...base,
        auditId: entry?.eventId || `${record?.screenCode}::closure::${index}`,
        module: "Billings",
        action:
          n(entry?.newStatus).toLowerCase() === "active"
            ? "Reactivated"
            : "Billing Status Changed",
        field: "Billing Status",
        previousValue: entry?.previousStatus,
        newValue: entry?.newStatus,
        changedBy: entry?.updatedBy || entry?.changedBy,
        changedAt:
          entry?.updatedAt || entry?.changedAt || entry?.effectiveDate,
        remarks:
          entry?.remarks || entry?.reason || entry?.restoreRemarks,
      });
    });

    const commercialHistory = Array.isArray(record?.commercialHistory)
      ? record.commercialHistory
      : [];

    commercialHistory.forEach((entry, index) => {
      addAudit(rows, {
        ...base,
        auditId:
          entry?.eventId || `${record?.screenCode}::commercial::${index}`,
        module: "Commercial",
        action: entry?.event || "Commercial Terms Changed",
        field: entry?.field || "Commercial Status",
        previousValue:
          entry?.previousValue ||
          entry?.previousStatus ||
          entry?.previousCommercialStatus,
        newValue:
          entry?.newValue ||
          entry?.newStatus ||
          entry?.commercialStatus,
        changedBy: entry?.changedBy || entry?.updatedBy,
        changedAt:
          entry?.changedAt || entry?.updatedAt || entry?.effectiveDate,
        remarks: entry?.remarks,
      });
    });

    if (record?.commercialUpdatedAt) {
      addAudit(rows, {
        ...base,
        module: "Commercial",
        action: "Commercial Record Updated",
        field: "Commercial Terms",
        previousValue: "",
        newValue:
          record?.commercialStatus ||
          record?.pricingMethod ||
          record?.billingMode ||
          "Updated",
        changedAt: record.commercialUpdatedAt,
        remarks: record?.addendumRemarks,
      });
    }

    if (record?.billingVerifiedAt) {
      addAudit(rows, {
        ...base,
        module: "Billings",
        action: "Billing Verification Updated",
        field: "Billing Verification Status",
        previousValue: "First Billing Pending Approval",
        newValue: record?.billingVerificationStatus,
        changedAt: record.billingVerifiedAt,
        remarks: record?.billingRemarks,
      });
    }
  });

  (billingRecords || []).forEach((record) => {
    const base = {
      billingCode: record?.billingCode,
      complexCode: record?.complexCode,
      screenCode: record?.screenCode || record?.siteScope,
      screenName: record?.screenName,
    };

    const entries = Array.isArray(record?.invoiceEntries)
      ? record.invoiceEntries
      : [];

    entries.forEach((entry, index) => {
      if (entry?.invoiceNumber || entry?.invoiceDate) {
        addAudit(rows, {
          ...base,
          auditId: `${entry?.entryId || index}::invoice`,
          module: "Billings",
          action: "Invoice Updated",
          field: "Invoice",
          previousValue: "",
          newValue: entry?.invoiceNumber || "Invoice recorded",
          changedBy: entry?.updatedBy,
          changedAt: entry?.updatedAt || entry?.invoiceDate,
          remarks: entry?.remarks,
        });
      }

      if (
        n(entry?.paymentStatus).toLowerCase() === "paid" ||
        entry?.receiptNumber ||
        entry?.paymentReceivedDate
      ) {
        addAudit(rows, {
          ...base,
          auditId: `${entry?.entryId || index}::payment`,
          module: "Billings",
          action: "Payment Updated",
          field: "Payment Status",
          previousValue: "Pending",
          newValue: entry?.paymentStatus || "Paid",
          changedBy: entry?.updatedBy,
          changedAt:
            entry?.paymentReceivedDate ||
            entry?.receiptDate ||
            entry?.updatedAt,
          remarks: entry?.paymentRemarks || entry?.remarks,
        });
      }
    });
  });

  Object.entries(incentiveStates || {}).forEach(([screenCode, state]) => {
    const stageRecord = (stage3Records || []).find(
      (record) => n(record?.screenCode) === n(screenCode),
    );

    const downloadedParts =
      state?.downloadedParts && typeof state.downloadedParts === "object"
        ? state.downloadedParts
        : {};

    Object.entries(downloadedParts).forEach(([partKey, part]) => {
      addAudit(rows, {
        billingCode: stageRecord?.billingCode,
        complexCode: stageRecord?.complexCode,
        screenCode,
        screenName: stageRecord?.screenName,
        module: "Incentive",
        action: "Incentive Downloaded",
        field: "Incentive Status",
        previousValue: "Ready",
        newValue: "Downloaded",
        changedAt: part?.downloadedAt,
        remarks: partKey,
      });
    });
  });

  governanceAuditRows.forEach((entry) => {
    const previous = parseJson(entry.previousValue);
    const next = parseJson(entry.newValue);
    const isCorrection = entry.action === "Record Correction";
    const parsedFields = parseJson(entry.field, null);
    const changedFields = isCorrection && Array.isArray(previous.changedFields)
      ? previous.changedFields
      : Array.isArray(parsedFields)
        ? parsedFields
        : [];
    const changedFieldText = Array.isArray(changedFields)
      ? changedFields.join(", ")
      : n(entry.field);
    const isCommercialUpdate =
      !isCorrection && entry.action === "Commercial Terms Updated";
    const isDedicatedStage3Movement =
      entry.module === "Workflow" && entry.action === "Moved to Stage 3";
    const isSnapshotStage3Movement =
      entry.module === "Site Entry" &&
      entry.action === "Stage 1 Site Updated" &&
      (n(previous.processingStatus).toLowerCase() === "stage 2" ||
        n(previous.readinessStatus).toLowerCase() === "moved to stage 2") &&
      (n(next.processingStatus).toLowerCase() === "stage 3" ||
        n(next.readinessStatus).toLowerCase() === "moved to stage 3");
    const isMovedToStage3 =
      isDedicatedStage3Movement || isSnapshotStage3Movement;
    const isDedicatedStage2Return =
      entry.module === "Workflow" && entry.action === "Returned to Stage 2";
    const isSnapshotStage2Return =
      entry.module === "Site Entry" &&
      entry.action === "Stage 1 Site Updated" &&
      (n(previous.processingStatus).toLowerCase() === "stage 3" ||
        n(previous.readinessStatus).toLowerCase() === "moved to stage 3") &&
      (n(next.processingStatus).toLowerCase() === "stage 2" ||
        n(next.readinessStatus).toLowerCase() === "moved to stage 2");
    const isReturnedToStage2 =
      isDedicatedStage2Return || isSnapshotStage2Return;
    const isStageMovement =
      isMovedToStage3 || isReturnedToStage2;

    addAudit(rows, {
      billingCode: entry.site?.billingId,
      complexCode: entry.site?.complexId,
      screenCode: entry.site?.siteId,
      screenName: entry.site?.screenName,
      module: isCorrection
        ? "Record Correction"
        : isStageMovement
          ? "Workflow"
          : isCommercialUpdate
            ? "Commercial"
            : entry.module,
      action: isCorrection
        ? "Record Correction"
        : isStageMovement
          ? isReturnedToStage2
            ? "Returned to Stage 2"
            : "Moved to Stage 3"
          : entry.action,
      workflowType: isCorrection
        ? "RECORD_CORRECTION"
        : isStageMovement
          ? "Stage Movement"
          : "",
      field: isStageMovement ? "Stage" : changedFieldText,
      previousValue: isCorrection
        ? previous?.previousValues?.processingStatus ||
          JSON.stringify(previous.previousValues || {})
        : isStageMovement
          ? isReturnedToStage2
            ? "Stage 3"
            : "Stage 2"
          : entry.previousValue,
      newValue: isCorrection
        ? next?.processingStatus === "Existing" &&
          next?.readinessStatus === "Ready"
          ? "Stage 1"
          : JSON.stringify(next || {})
        : isStageMovement
          ? isReturnedToStage2
            ? "Stage 2"
            : "Stage 3"
          : entry.newValue,
      rawPreviousValue: isCorrection
        ? JSON.stringify(previous.previousValues || {})
        : isStageMovement
          ? isReturnedToStage2
            ? "Stage 3"
            : "Stage 2"
          : entry.previousValue,
      rawNewValue: isStageMovement
        ? isReturnedToStage2
          ? "Stage 2"
          : "Stage 3"
        : entry.newValue,
      changedBy: entry.changedBy,
      changedAt: entry.createdAt,
      remarks: entry.remarks,
      changedFields: isCorrection || isCommercialUpdate ? changedFields : [],
      auditId: entry.id,
    });
  });

  const deduplicatedRows = governanceAuditRows.length
    ? rows.filter((row) => row.action !== "Commercial Record Updated")
    : rows;

  return deduplicatedRows.sort((a, b) => {
    const aTime = parseDate(a.changedAt)?.getTime() || 0;
    const bTime = parseDate(b.changedAt)?.getTime() || 0;
    return bTime - aTime;
  });
}

export default function AuditHistory({
  stage1Records = [],
  stage2Records = [],
  stage3Records = [],
  billingRecords = [],
  incentiveStates = {},
}) {
  const [search, setSearch] = useState("");
  const [moduleFilter, setModuleFilter] = useState("All");
  const [actionFilter, setActionFilter] = useState("All");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [activeAudit, setActiveAudit] = useState(null);
  const [rowsPerPage, setRowsPerPage] = useState(10);
  const [currentPage, setCurrentPage] = useState(1);
  const [governanceAuditRows, setGovernanceAuditRows] = useState([]);

  useEffect(() => {
    let cancelled = false;
    billingApiRequest("/audit-history")
      .then((records) => {
        if (!cancelled) setGovernanceAuditRows(Array.isArray(records) ? records : []);
      })
      .catch((error) => {
        console.error("Unable to load Governance audit history:", error);
        if (!cancelled) setGovernanceAuditRows([]);
      });

    return () => {
      cancelled = true;
    };
  }, []);

  const rows = useMemo(
    () =>
      collectAuditRows({
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

  const filtered = useMemo(() => {
    const q = n(search).toLowerCase();
    const fromDate = dateFrom ? new Date(`${dateFrom}T00:00:00`) : null;
    const toDate = dateTo ? new Date(`${dateTo}T23:59:59`) : null;

    return rows.filter((row) => {
      if (moduleFilter !== "All" && n(row.module) !== moduleFilter) return false;
      if (
        actionFilter !== "All" &&
        getGovernanceEventCategory(row.action) !== actionFilter
      ) return false;
      const rowDate = parseDate(row.changedAt);
      if (fromDate && rowDate && rowDate < fromDate) return false;
      if (toDate && rowDate && rowDate > toDate) return false;

      if (!q) return true;

      return [
        row.billingCode,
        row.complexCode,
        row.screenCode,
        row.screenName,
        row.module,
        row.action,
        row.workflowType,
        row.field,
        row.previousValue,
        row.newValue,
        row.changedBy,
        row.remarks,
      ]
        .join(" ")
        .toLowerCase()
        .includes(q);
    });
  }, [rows, search, moduleFilter, actionFilter, dateFrom, dateTo]);

  const moduleOptions = useMemo(
    () => sortedFilterOptions(rows.map((row) => row.module)),
    [rows],
  );
  const actionOptions = GOVERNANCE_EVENT_CATEGORIES;

  useEffect(() => {
    setCurrentPage(1);
  }, [search, moduleFilter, actionFilter, dateFrom, dateTo, rowsPerPage]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / rowsPerPage));

  useEffect(() => {
    setCurrentPage((page) => Math.min(page, totalPages));
  }, [totalPages]);

  const pagedRows = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filtered.slice(start, start + rowsPerPage);
  }, [currentPage, filtered, rowsPerPage]);

  const firstVisible = filtered.length
    ? (currentPage - 1) * rowsPerPage + 1
    : 0;
  const lastVisible = Math.min(currentPage * rowsPerPage, filtered.length);

  return (
    <div className="governance-history">
      <div style={{ marginBottom: 18 }}>
        <div className="governance-history__eyebrow">
          Governance
        </div>
        <h1>Audit History</h1>
        <p className="governance-history__intro">
          Read-only audit trail of operational, billing, commercial, and
          incentive changes. Most recent change is shown first.
        </p>
      </div>

      <div className="governance-history__filters governance-history__filters--audit">
        <input
          className="governance-history__control"
          type="search"
          placeholder="Search Screen, Billing ID, field, action, user..."
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

        <select className="governance-history__control" value={actionFilter} onChange={(event) => setActionFilter(event.target.value)} aria-label="Filter by event">
          <option value="All">All Events</option>
          {actionOptions.map((category) => <option key={category} value={category}>{category}</option>)}
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

        <button type="button" className="governance-history__button" onClick={() => {
          setSearch("");
          setModuleFilter("All");
          setActionFilter("All");
          setDateFrom("");
          setDateTo("");
          setCurrentPage(1);
        }}>Reset</button>
      </div>

      <div className="governance-history__card">
        <div className="governance-history__card-header">
          <strong>Audit Records</strong>
          <span className="governance-history__count">{filtered.length} records</span>
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
                  "Action",
                  "View",
                ].map((heading, index) => (
                  <th key={`${heading}-${index}`}>{heading}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length ? (
                pagedRows.map((row) => (
                  <tr key={row.auditId}>
                    <td>{formatDateTime(row.changedAt)}</td>
                    <td><span className="governance-history__ellipsis" title={row.screenCode}>{row.screenCode || "-"}</span></td>
                    <td><span className="governance-history__ellipsis" title={row.module}>{row.module || "-"}</span></td>
                    <td><span className="governance-history__ellipsis" title={compactAuditAction(row.action)}>{compactAuditAction(row.action) || "-"}</span></td>
                    <td><button type="button" className="governance-history__button governance-history__view" onClick={() => setActiveAudit(row)}>View</button></td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={5} className="governance-history__empty">
                    No audit records found.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        {filtered.length ? (
          <div className="governance-history__pagination">
            <div className="governance-history__pagination-summary">
              <span>Showing {firstVisible}-{lastVisible} of {filtered.length}</span>
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

      {activeAudit ? (
        <div className="governance-history__backdrop" role="presentation" onMouseDown={() => setActiveAudit(null)}>
          <section
            className="governance-history__dialog"
            role="dialog"
            aria-modal="true"
            aria-labelledby="audit-detail-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="governance-history__dialog-header">
              <div>
                <div className="governance-history__screen">{activeAudit.screenCode}</div>
                <h2 id="audit-detail-title">{activeAudit.action}</h2>
              </div>
              <button type="button" className="governance-history__button" onClick={() => setActiveAudit(null)}>Close</button>
            </div>

            <div className="governance-history__metadata">
              <div><span>Screen Code</span><br />{activeAudit.screenCode || "-"}</div>
              <div><span>Module</span><br />{activeAudit.module || "-"}</div>
              <div><span>Action</span><br />{compactAuditAction(activeAudit.action) || "-"}</div>
              <div><span>Billing Code</span><br />{activeAudit.billingCode || "-"}</div>
              <div><span>Complex Code</span><br />{activeAudit.complexCode || "Standalone"}</div>
              <div><span>Screen Name</span><br />{activeAudit.screenName || "-"}</div>
              <div><span>Save Mode</span><br />{activeAudit.saveMode || "Not Available"}</div>
              <div><span>Reason</span><br />{activeAudit.remarks || "-"}</div>
              <div><span>Changed By</span><br />{activeAudit.changedBy || "Not Available"}</div>
              <div><span>Date &amp; Time</span><br />{formatDateTime(activeAudit.changedAt)}</div>
              <div><span>Workflow Type</span><br />{activeAudit.workflowType === "RECORD_CORRECTION" ? "Record Correction" : activeAudit.workflowType || "-"}</div>
            </div>

            <div className="governance-history__detail-heading">Changed Fields</div>
            {detailChanges(activeAudit).length ? (
              <table className="governance-history__detail-table">
                <thead><tr><th>Field</th><th>Previous</th><th>New</th></tr></thead>
                <tbody>
                  {detailChanges(activeAudit).map((change, index) => (
                    <tr key={`${change.field}-${index}`}>
                      <td>{change.label}</td>
                      <td>{formatDetailValue(change.field, change.previous)}</td>
                      <td>{formatDetailValue(change.field, change.next)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <div className="governance-history__detail-empty">No field-level changes recorded</div>
            )}
          </section>
        </div>
      ) : null}
    </div>
  );
}
