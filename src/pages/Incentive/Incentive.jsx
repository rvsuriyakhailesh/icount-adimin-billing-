import { useEffect, useMemo, useRef, useState } from "react";
import { billingApiRequest } from "../../utils/billingApi";
import "./Incentive.css";

function n(value) {
  return String(value ?? "").trim();
}

function compareColumnValues(leftValue, rightValue) {
  const leftText = n(leftValue);
  const rightText = n(rightValue);

  const leftDate = /^\d{4}-\d{2}-\d{2}/.test(leftText)
    ? Date.parse(leftText)
    : NaN;
  const rightDate = /^\d{4}-\d{2}-\d{2}/.test(rightText)
    ? Date.parse(rightText)
    : NaN;

  if (!Number.isNaN(leftDate) && !Number.isNaN(rightDate)) {
    return leftDate - rightDate;
  }

  const leftNumber = Number(leftText.replace(/[^0-9.-]/g, ""));
  const rightNumber = Number(rightText.replace(/[^0-9.-]/g, ""));

  if (
    leftText &&
    rightText &&
    Number.isFinite(leftNumber) &&
    Number.isFinite(rightNumber)
  ) {
    return leftNumber - rightNumber;
  }

  return leftText.localeCompare(rightText, undefined, {
    numeric: true,
    sensitivity: "base",
  });
}

function SortableHeader({
  label,
  sortKey,
  activeKey,
  direction,
  onSort,
}) {
  const active = activeKey === sortKey;

  return (
    <th>
      <button
        type="button"
        onClick={() => onSort(sortKey)}
        className="incentive-page__sort-button"
        title={`Sort ${label}`}
      >
        {label} {active ? (direction === "asc" ? "▲" : "▼") : ""}
      </button>
    </th>
  );
}

function parseDisplayDate(value) {
  const text = n(value).slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;

  const [year, month, day] = text.split("-").map(Number);
  const date = new Date(year, month - 1, day);

  return Number.isNaN(date.getTime()) ? null : date;
}

function displayDate(value) {
  const date = parseDisplayDate(value);
  return date
    ? date.toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "short",
        year: "numeric",
      })
    : n(value) || "-";
}

function money(value) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 0,
  }).format(Number(value || 0));
}

function todayInputValue() {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, "0");
  const day = String(today.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

const PAYMENT_REMARKS_MAX_LENGTH = 250;

function billingMonthKey(value) {
  const text = n(value).slice(0, 7);
  return /^\d{4}-\d{2}$/.test(text) ? text : "";
}

function billingEntryCoversRequiredMonth(entry, requiredMonth) {
  const fromMonth = billingMonthKey(entry?.periodFrom);
  const toMonth = billingMonthKey(entry?.periodTo);

  return Boolean(
    fromMonth &&
      toMonth &&
      requiredMonth >= fromMonth &&
      requiredMonth <= toMonth,
  );
}

function isReceivedBillingEvidence(entry) {
  const receivedStatuses = [
    "paid",
    "payment received",
    "received",
    "completed",
    "complete",
  ];
  const status = n(entry?.paymentStatus).toLowerCase();
  const receivedDate = parseDisplayDate(
    entry?.paymentReceivedDate || entry?.receiptDate,
  );
  const today = parseDisplayDate(todayInputValue());

  return Boolean(
    receivedStatuses.includes(status) &&
      receivedDate &&
      today &&
      receivedDate <= today,
  );
}

function billingEvidenceEntries(record) {
  const entries = Array.isArray(record?.invoiceEntries)
    ? record.invoiceEntries
    : [];
  const sourceEntries = entries.length > 0 ? entries : [{}];

  return sourceEntries.map((entry) => ({
    periodFrom: n(entry?.periodFrom || record?.billingPeriodFrom),
    periodTo: n(entry?.periodTo || record?.billingPeriodTo),
    invoiceNumber: n(entry?.invoiceNumber || record?.invoiceNumber),
    paymentStatus: n(entry?.paymentStatus || record?.paymentStatus),
    paymentReceivedDate: n(
      entry?.paymentReceivedDate ||
        entry?.receiptDate ||
        record?.paymentReceivedDate ||
        record?.receiptDate ||
        record?.paymentReceiptDate ||
        record?.erpReceiptDate,
    ),
    receiptNumber: n(
      entry?.receiptNumber ||
        record?.receiptNumber ||
        record?.paymentReceiptNumber ||
        record?.erpReceiptNumber,
    ),
    receiptDate: n(
      entry?.receiptDate ||
        record?.receiptDate ||
        record?.paymentReceiptDate ||
        record?.erpReceiptDate,
    ),
    invoiceAmountBeforeGST: n(
      entry?.invoiceAmountBeforeGST ||
        record?.invoiceAmountBeforeGST ||
        record?.billingAmountBeforeGST,
    ),
  }));
}

function qualifyingBillingEvidence(record, requiredMonths) {
  return billingEvidenceEntries(record).filter((entry) => {
    if (!isReceivedBillingEvidence(entry)) return false;
    if (requiredMonths.length === 0) return true;

    return requiredMonths.some((month) =>
      billingEntryCoversRequiredMonth(entry, month),
    );
  });
}

function billingRecordScreenCode(record) {
  return n(record?.screenCode || record?.siteScope).toUpperCase();
}

function commonBillingGroupKey(record) {
  return n(
    record?.combinedGroupKey ||
      record?.stage1GroupId ||
      record?.complexCode ||
      record?.billingCode,
  ).toUpperCase();
}

function isCommonBillingRecord(record) {
  return n(record?.pricingMethod).toLowerCase() === "common";
}

function isFoCBillingRecord(record) {
  return (
    n(record?.originStage).toLowerCase() === "foc" ||
    n(record?.commercialType).toLowerCase() === "foc" ||
    n(record?.commercialApplicable).toLowerCase() === "false" ||
    n(record?.focStatus).toLowerCase() === "foc"
  );
}

function billingRecordEffectiveScreenFee(record) {
  const screenCode = billingRecordScreenCode(record);
  const commonRows = Array.isArray(record?.commonComplexBillingRows)
    ? record.commonComplexBillingRows
    : [];
  const commonRow = commonRows.find(
    (row) => billingRecordScreenCode(row) === screenCode,
  );

  return n(
    record?.effectiveScreenFee ||
      commonRow?.allocatedFee ||
      record?.billingCommercialLockSnapshot?.subscriptionFee ||
      record?.actualSubscriptionFee ||
      record?.subscriptionFee ||
      record?.billingAmountBeforeGST,
  );
}

function numericAmount(value) {
  const amount = Number(n(value).replace(/[^0-9.-]/g, ""));
  return Number.isFinite(amount) ? amount : 0;
}

function uniqueValues(values) {
  return Array.from(new Set(values.map(n).filter(Boolean)));
}

function paymentScopeMatches(entry, invoiceNumbers, receiptNumbers, periods) {
  const invoiceNumber = n(entry.invoiceNumber);
  const receiptNumber = n(entry.receiptNumber);
  const hasDocumentScope = invoiceNumbers.size > 0 || receiptNumbers.size > 0;

  if (hasDocumentScope) {
    return (
      (invoiceNumber && invoiceNumbers.has(invoiceNumber)) ||
      (receiptNumber && receiptNumbers.has(receiptNumber))
    );
  }

  return periods.has(`${n(entry.periodFrom)}::${n(entry.periodTo)}`);
}

function generatedPaymentRemark(row, milestone, stage3Records, billingRecords) {
  const screenCode = n(row?.screenCode).toUpperCase();
  const requiredMonths = Array.isArray(milestone?.requiredKeys)
    ? milestone.requiredKeys.map(n).filter(Boolean)
    : [];
  const sourceRecord = stage3Records.find(
    (record) => n(record?.screenCode).toUpperCase() === screenCode,
  );
  const sourceRecordId = n(sourceRecord?.recordId || row?.recordId);
  const targetBillingRecords = billingRecords.filter(
    (record) =>
      billingRecordScreenCode(record) === screenCode ||
      (sourceRecordId && n(record?.sourceRecordId) === sourceRecordId),
  );
  const targetEvidence = targetBillingRecords.flatMap((record) =>
    qualifyingBillingEvidence(record, requiredMonths),
  );

  if (targetEvidence.length === 0) {
    return "";
  }

  const commonRecord =
    [sourceRecord, ...targetBillingRecords].find(isCommonBillingRecord) || null;
  const isCommon = Boolean(commonRecord);
  const groupKey = isCommon ? commonBillingGroupKey(commonRecord) : "";
  const targetInvoiceNumbers = new Set(
    uniqueValues(targetEvidence.map((entry) => entry.invoiceNumber)),
  );
  const targetReceiptNumbers = new Set(
    uniqueValues(targetEvidence.map((entry) => entry.receiptNumber)),
  );
  const targetPeriods = new Set(
    targetEvidence.map(
      (entry) => `${n(entry.periodFrom)}::${n(entry.periodTo)}`,
    ),
  );
  const participatingRecords = isCommon
    ? billingRecords.filter(
        (record) =>
          !record?.isCombinedCommonBilling &&
          !isFoCBillingRecord(record) &&
          isCommonBillingRecord(record) &&
          commonBillingGroupKey(record) === groupKey &&
          Boolean(billingRecordScreenCode(record)),
      )
    : targetBillingRecords;
  const participants = new Map();
  const matchedEvidence = [];

  participatingRecords.forEach((record) => {
    const evidence = qualifyingBillingEvidence(record, requiredMonths).filter(
      (entry) =>
        !isCommon ||
        paymentScopeMatches(
          entry,
          targetInvoiceNumbers,
          targetReceiptNumbers,
          targetPeriods,
        ),
    );

    if (evidence.length === 0) return;

    const participantScreen = billingRecordScreenCode(record);
    if (!participantScreen) return;

    const effectiveFee =
      billingRecordEffectiveScreenFee(record) ||
      evidence.find((entry) => entry.invoiceAmountBeforeGST)
        ?.invoiceAmountBeforeGST ||
      "";

    if (!participants.has(participantScreen)) {
      participants.set(participantScreen, {
        screenCode: participantScreen,
        effectiveFee,
      });
    }
    matchedEvidence.push(...evidence);
  });

  if (participants.size === 0) {
    return "";
  }

  const participantRows = Array.from(participants.values()).sort((left, right) =>
    left.screenCode.localeCompare(right.screenCode, undefined, {
      numeric: true,
    }),
  );
  const invoiceNumbers = uniqueValues(
    matchedEvidence.map((entry) => entry.invoiceNumber),
  );
  const receiptNumbers = uniqueValues(
    matchedEvidence.map((entry) => entry.receiptNumber),
  );
  const amount = participantRows.reduce(
    (total, participant) => total + numericAmount(participant.effectiveFee),
    0,
  );
  const segments = isCommon
    ? [
        participantRows.map((participant) => participant.screenCode).join(", "),
        `Combined ${money(amount)}`,
      ]
    : [participantRows[0].screenCode, money(amount)];

  if (invoiceNumbers.length > 0) {
    segments.push(`Inv: ${invoiceNumbers.join(" / ")}`);
  }
  if (receiptNumbers.length > 0) {
    segments.push(`Rcpt: ${receiptNumbers.join(" / ")}`);
  }

  const remark = segments.join(" | ");
  return remark.length <= PAYMENT_REMARKS_MAX_LENGTH
    ? remark
    : `${remark.slice(0, PAYMENT_REMARKS_MAX_LENGTH - 3).trimEnd()}...`;
}

function buildLegacyIncentiveState(rows) {
  const next = {};

  rows.forEach((row) => {
    const downloadedParts = {};

    (row?.milestones || []).forEach((milestone) => {
      if (!milestone?.downloadedAt && !milestone?.paidDate) return;

      downloadedParts[milestone.key] = {
        downloadedAt: milestone.downloadedAt || milestone.paidDate,
        paidDate: milestone.paidDate || "",
        paymentRemarks: milestone.paymentRemarks || "",
        signature: row.signature || "",
        amount: milestone.amount,
        beneficiary: row.beneficiary || "",
        backendStatus: milestone.status,
      };
    });

    next[n(row.screenCode)] = {
      backendIncentiveId: row.backendIncentiveId,
      currentStatus: row.currentStatus,
      downloadedParts,
    };
  });

  return next;
}

function partTone(status) {
  const value = n(status).toLowerCase();

  if (value === "paid") return "paid";
  if (value === "ready" || value === "downloaded") return "ready";
  if (value === "review required") return "review";
  return "not-ready";
}

function partTooltip(milestone) {
  if (!milestone) return "No Incentive part";

  if (milestone.status === "Paid") {
    return `Paid${milestone.paidDate ? ` on ${displayDate(milestone.paidDate)}` : ""}. Click to view payment details.`;
  }

  if (["Ready", "Downloaded"].includes(milestone.status)) {
    return "Ready for Domestic team payment. Click to record Payment Date and Remarks.";
  }

  if (milestone.status === "Review Required") {
    return "Review Required. Open View Details for the reason.";
  }

  return milestone.missing?.length
    ? milestone.missing.join(" | ")
    : "Not Ready";
}

function Incentive({
  stage3Records = [],
  billingRecords = [],
  setIncentiveStates = () => {},
}) {
  const [rows, setRows] = useState([]);
  const [search, setSearch] = useState("");
  const [sortOrder, setSortOrder] = useState("desc");
  const [sortKey, setSortKey] = useState("");
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const [activeScreen, setActiveScreen] = useState("");
  const [loading, setLoading] = useState(true);
  const [apiError, setApiError] = useState("");
  const [lastEvaluatedAt, setLastEvaluatedAt] = useState("");
  const [rowsPerPage, setRowsPerPage] = useState(10);
  const [currentPage, setCurrentPage] = useState(1);

  const [paymentTarget, setPaymentTarget] = useState(null);
  const [paymentDate, setPaymentDate] = useState("");
  const [paymentRemarks, setPaymentRemarks] = useState("");
  const [paymentBusy, setPaymentBusy] = useState(false);
  const [successMessage, setSuccessMessage] = useState("");
  const incentivePageRef = useRef(null);
  const [isFullscreen, setIsFullscreen] = useState(false);

  useEffect(() => {
    function handleFullscreenChange() {
      setIsFullscreen(
        Boolean(
          incentivePageRef.current &&
            document.fullscreenElement === incentivePageRef.current,
        ),
      );
    }

    document.addEventListener("fullscreenchange", handleFullscreenChange);
    return () => {
      document.removeEventListener(
        "fullscreenchange",
        handleFullscreenChange,
      );
    };
  }, []);

  async function toggleIncentiveFullscreen() {
    try {
      if (
        incentivePageRef.current &&
        document.fullscreenElement !== incentivePageRef.current
      ) {
        await incentivePageRef.current.requestFullscreen();
      } else if (document.fullscreenElement) {
        await document.exitFullscreen();
      }
    } catch (error) {
      setApiError(
        error?.message || "Unable to change Incentive fullscreen mode.",
      );
    }
  }

  function applyBackendRows(nextRows) {
    const normalized = Array.isArray(nextRows) ? nextRows : [];
    setRows(normalized);

    const legacyMap = buildLegacyIncentiveState(normalized);
    setIncentiveStates((current) => ({
      ...(current && typeof current === "object" ? current : {}),
      ...legacyMap,
    }));
  }

  async function evaluateAndRefresh() {
    setLoading(true);
    setApiError("");

    try {
      const data = await billingApiRequest("/incentives/evaluate", {
        method: "POST",
        body: JSON.stringify({
          billingRecords: Array.isArray(billingRecords) ? billingRecords : [],
          changedBy: "Domestic Team",
        }),
      });

      applyBackendRows(data);
      setLastEvaluatedAt(new Date().toISOString());
    } catch (error) {
      setApiError(
        error?.message ||
          "Unable to evaluate Incentive records from the backend.",
      );

      try {
        const fallback = await billingApiRequest("/incentives");
        applyBackendRows(fallback);
      } catch {
        // Keep the evaluation error visible.
      }
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    let cancelled = false;

    async function load() {
      setLoading(true);
      setApiError("");

      try {
        const data = await billingApiRequest("/incentives");

        if (cancelled) return;

        applyBackendRows(data);
        setLastEvaluatedAt(new Date().toISOString());
      } catch (error) {
        if (cancelled) return;

        setApiError(
          error?.message ||
            "Unable to evaluate Incentive records from the backend.",
        );

        try {
          const fallback = await billingApiRequest("/incentives");
          if (!cancelled) applyBackendRows(fallback);
        } catch {
          // Keep the evaluation error.
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    }

    load();

    return () => {
      cancelled = true;
    };
  }, [stage3Records, billingRecords]);

  const filtered = useMemo(() => {
    const query = n(search).toLowerCase();
    const from = fromDate ? new Date(`${fromDate}T00:00:00`) : null;
    const to = toDate ? new Date(`${toDate}T23:59:59`) : null;

    return rows
      .filter((row) => {
        if (
          query &&
          ![
            row.billingCode,
            row.complexCode,
            row.screenCode,
            row.screenName,
            row.location,
            row.state,
            row.beneficiary,
            row.employeeNumber,
            row.companyId,
          ]
            .join(" ")
            .toLowerCase()
            .includes(query)
        ) {
          return false;
        }

        if (from || to) {
          const eligibilityValue =
            (row.milestones || []).find((milestone) =>
              n(milestone.eligibilityDate),
            )?.eligibilityDate || row.billingStartDate;

          const eligibilityDate = eligibilityValue
            ? new Date(`${String(eligibilityValue).slice(0, 10)}T12:00:00`)
            : null;

          if (!eligibilityDate || Number.isNaN(eligibilityDate.getTime())) {
            return false;
          }

          if (from && eligibilityDate < from) return false;
          if (to && eligibilityDate > to) return false;
        }

        return true;
      })
      .sort((left, right) => {
        const valueFor = (row) => {
          if (!sortKey) {
            return (
              (row.milestones || []).find((milestone) =>
                n(milestone.eligibilityDate),
              )?.eligibilityDate ||
              row.billingStartDate ||
              ""
            );
          }

          if (sortKey.startsWith("part")) {
            const index = Number(sortKey.slice(4)) - 1;
            const milestone = row.milestones?.[index];
            return milestone ? `${milestone.status} ${milestone.amount}` : "";
          }

          return row?.[sortKey];
        };

        const comparison = compareColumnValues(
          valueFor(left),
          valueFor(right),
        );

        return sortOrder === "asc" ? comparison : -comparison;
      });
  }, [rows, search, sortOrder, sortKey, fromDate, toDate]);

  useEffect(() => {
    setCurrentPage(1);
  }, [search, fromDate, toDate, rowsPerPage]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / rowsPerPage));

  useEffect(() => {
    if (currentPage > totalPages) setCurrentPage(totalPages);
  }, [currentPage, totalPages]);

  const pageRows = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filtered.slice(start, start + rowsPerPage);
  }, [filtered, currentPage, rowsPerPage]);

  const activeRow =
    rows.find((row) => n(row.screenCode) === activeScreen) || null;

  function handleColumnSort(key) {
    if (sortKey === key) {
      setSortOrder((current) => (current === "asc" ? "desc" : "asc"));
      return;
    }

    setSortKey(key);
    setSortOrder("asc");
  }

  function openPayment(row, milestone) {
    if (!milestone) return;

    const payable = ["Ready", "Downloaded"].includes(milestone.status);
    const paid = milestone.status === "Paid";

    if (!payable && !paid) return;

    setPaymentTarget({ row, milestone });
    setPaymentDate(milestone.status === "Paid" ? milestone.paidDate || "" : "");
    setPaymentRemarks(
      paid
        ? milestone.paymentRemarks || ""
        : generatedPaymentRemark(
            row,
            milestone,
            stage3Records,
            billingRecords,
          ),
    );
    setApiError("");
    setSuccessMessage("");
  }

  function clearPaymentDraft() {
    setPaymentTarget(null);
    setPaymentDate("");
    setPaymentRemarks("");
  }

  function closePayment() {
    if (paymentBusy) return;
    clearPaymentDraft();
  }

  async function savePayment() {
    if (paymentBusy) return;

    if (!paymentTarget?.row || !paymentTarget?.milestone) {
      setApiError("Select a Ready Incentive Part before saving payment.");
      return;
    }

    if (
      !["Ready", "Downloaded"].includes(paymentTarget.milestone.status)
    ) {
      setApiError("Only a Ready Incentive Part can be marked Paid.");
      return;
    }

    if (!paymentDate) {
      setApiError("Payment Date is required.");
      return;
    }

    const normalizedPaymentRemarks = n(paymentRemarks);

    if (!normalizedPaymentRemarks) {
      setApiError("Payment Remarks is required.");
      return;
    }

    if (normalizedPaymentRemarks.length > PAYMENT_REMARKS_MAX_LENGTH) {
      setApiError(
        `Payment Remarks must be ${PAYMENT_REMARKS_MAX_LENGTH} characters or fewer.`,
      );
      return;
    }

    const screenCode = n(paymentTarget.row.screenCode);
    const milestoneKey = n(paymentTarget.milestone.key);
    const normalizedPaymentDate = n(paymentDate);
    const parsedPaymentDate = parseDisplayDate(normalizedPaymentDate);
    const eligibilityDate = parseDisplayDate(
      paymentTarget.milestone.eligibilityDate,
    );
    const today = parseDisplayDate(todayInputValue());

    if (!screenCode || !/^part[123]$/.test(milestoneKey)) {
      setApiError("The selected Incentive payment record is invalid.");
      return;
    }

    if (
      !/^\d{4}-\d{2}-\d{2}$/.test(normalizedPaymentDate) ||
      !parsedPaymentDate
    ) {
      setApiError("Payment Date must use YYYY-MM-DD.");
      return;
    }

    if (eligibilityDate && parsedPaymentDate < eligibilityDate) {
      setApiError("Payment Date cannot be earlier than the Incentive Eligibility Date.");
      return;
    }

    if (today && parsedPaymentDate > today) {
      setApiError("Payment Date cannot be in the future.");
      return;
    }

    setPaymentBusy(true);
    setApiError("");
    setSuccessMessage("");

    try {
      const updatedRows = await billingApiRequest("/incentives/mark-paid", {
        method: "POST",
        body: JSON.stringify({
          screenCode,
          milestoneKey,
          paymentDate: normalizedPaymentDate,
          remarks: normalizedPaymentRemarks,
          paidBy: "Domestic Team",
        }),
      });

      applyBackendRows(updatedRows);
      clearPaymentDraft();
      setSuccessMessage("Incentive payment saved successfully.");
    } catch (error) {
      setApiError(
        error?.message || "Unable to save the Incentive payment.",
      );
    } finally {
      setPaymentBusy(false);
    }
  }

  const firstVisible =
    filtered.length === 0 ? 0 : (currentPage - 1) * rowsPerPage + 1;
  const lastVisible = Math.min(currentPage * rowsPerPage, filtered.length);

  return (
    <section
      ref={incentivePageRef}
      className={`incentive-page${isFullscreen ? " incentive-page--fullscreen" : ""}`}
    >
      <div className="incentive-page__header">
        <div>
          <p className="incentive-page__eyebrow">Incentive</p>
          <h1>Incentive Eligibility &amp; Payment</h1>
          <p className="incentive-page__subtitle">
            Incentive eligibility is calculated by the backend from Billing
            Start Date, the applicable Incentive block and actual client
            payment receipt. When a Part becomes Ready, select its amount to
            record the Domestic team Payment Date and Remarks.
          </p>
        </div>

        <div className="incentive-page__header-actions">
          <button
            type="button"
            className="incentive-page__secondary-button"
            onClick={evaluateAndRefresh}
            disabled={loading || paymentBusy}
          >
            {loading ? "Evaluating..." : "Refresh Eligibility"}
          </button>

          <button
            type="button"
            className="incentive-page__secondary-button incentive-page__fullscreen-button"
            onClick={toggleIncentiveFullscreen}
            title={isFullscreen ? "Exit Full Screen" : "Full Screen"}
            aria-label={isFullscreen ? "Exit Incentive full screen" : "Open Incentive full screen"}
          >
            {isFullscreen ? (
              <svg
                width="18"
                height="18"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
                aria-hidden="true"
              >
                <path d="M9 3v6H3" />
                <path d="m3 9 6-6" />
                <path d="M15 3v6h6" />
                <path d="m21 9-6-6" />
                <path d="M9 21v-6H3" />
                <path d="m3 15 6 6" />
                <path d="M15 21v-6h6" />
                <path d="m21 15-6 6" />
              </svg>
            ) : (
              <svg
                width="18"
                height="18"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
                aria-hidden="true"
              >
                <path d="M8 3H3v5" />
                <path d="M3 3l6 6" />
                <path d="M16 3h5v5" />
                <path d="m21 3-6 6" />
                <path d="M8 21H3v-5" />
                <path d="m3 21 6-6" />
                <path d="M16 21h5v-5" />
                <path d="m21 21-6-6" />
              </svg>
            )}
          </button>
        </div>
      </div>

      {apiError && !paymentTarget ? (
        <div className="incentive-page__api-message incentive-page__api-message--error">
          <strong>Incentive:</strong> {apiError}
        </div>
      ) : null}

      {successMessage && !apiError ? (
        <div className="incentive-page__api-message incentive-page__api-message--success">
          {successMessage}
        </div>
      ) : null}

      {lastEvaluatedAt && !apiError && !successMessage ? (
        <div className="incentive-page__api-message incentive-page__api-message--success">
          Incentive eligibility loaded from PostgreSQL/backend.
        </div>
      ) : null}

      <article className="incentive-page__card">
        <div className="incentive-page__toolbar">
          <div className="incentive-page__field incentive-page__field--search">
            <label>Search</label>
            <input
              type="search"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Search Screen, Location, State, Beneficiary or Company ID"
            />
          </div>

          <div className="incentive-page__field">
            <label>From Date</label>
            <input
              type="date"
              value={fromDate}
              onChange={(event) => setFromDate(event.target.value)}
            />
          </div>

          <div className="incentive-page__field">
            <label>To Date</label>
            <input
              type="date"
              value={toDate}
              onChange={(event) => setToDate(event.target.value)}
            />
          </div>

          <div className="incentive-page__field incentive-page__field--rpp">
            <label>RPP</label>
            <select
              value={rowsPerPage}
              onChange={(event) => setRowsPerPage(Number(event.target.value))}
            >
              {[10, 20, 30, 50].map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </div>
        </div>

        {loading && rows.length === 0 ? (
          <div className="incentive-page__empty">
            Evaluating Incentive eligibility...
          </div>
        ) : filtered.length === 0 ? (
          <div className="incentive-page__empty">
            No Incentive records are available yet.
          </div>
        ) : (
          <>
            <div className="incentive-page__table-wrapper">
              <table className="incentive-page__table">
                <colgroup>
                  <col style={{ width: "112px" }} />
                  <col />
                  <col />
                  <col style={{ width: "100px" }} />
                  <col />
                  <col style={{ width: "110px" }} />
                  <col style={{ width: "80px" }} />
                  <col style={{ width: "80px" }} />
                  <col style={{ width: "80px" }} />
                  <col style={{ width: "56px" }} />
                </colgroup>
                <thead>
                  <tr>
                    <SortableHeader
                      label="Screen Code"
                      sortKey="screenCode"
                      activeKey={sortKey}
                      direction={sortOrder}
                      onSort={handleColumnSort}
                    />
                    <SortableHeader
                      label="Screen Name"
                      sortKey="screenName"
                      activeKey={sortKey}
                      direction={sortOrder}
                      onSort={handleColumnSort}
                    />
                    <SortableHeader
                      label="Location"
                      sortKey="location"
                      activeKey={sortKey}
                      direction={sortOrder}
                      onSort={handleColumnSort}
                    />
                    <SortableHeader
                      label="State"
                      sortKey="state"
                      activeKey={sortKey}
                      direction={sortOrder}
                      onSort={handleColumnSort}
                    />
                    <SortableHeader
                      label="Beneficiary"
                      sortKey="beneficiary"
                      activeKey={sortKey}
                      direction={sortOrder}
                      onSort={handleColumnSort}
                    />
                    <SortableHeader
                      label="Company ID"
                      sortKey="employeeNumber"
                      activeKey={sortKey}
                      direction={sortOrder}
                      onSort={handleColumnSort}
                    />
                    <SortableHeader
                      label="Part 1"
                      sortKey="part1"
                      activeKey={sortKey}
                      direction={sortOrder}
                      onSort={handleColumnSort}
                    />
                    <SortableHeader
                      label="Part 2"
                      sortKey="part2"
                      activeKey={sortKey}
                      direction={sortOrder}
                      onSort={handleColumnSort}
                    />
                    <SortableHeader
                      label="Part 3"
                      sortKey="part3"
                      activeKey={sortKey}
                      direction={sortOrder}
                      onSort={handleColumnSort}
                    />
                    <th>Action</th>
                  </tr>
                </thead>

                <tbody>
                  {pageRows.map((row) => (
                    <tr key={row.screenCode}>
                      <td title={row.screenCode || undefined}><span className="table-cell-ellipsis">{row.screenCode || "-"}</span></td>
                      <td title={row.screenName || undefined}><span className="table-cell-ellipsis">{row.screenName || "-"}</span></td>
                      <td title={row.location || undefined}><span className="table-cell-ellipsis">{row.location || "-"}</span></td>
                      <td title={row.state || undefined}><span className="table-cell-ellipsis">{row.state || "-"}</span></td>
                      <td title={row.beneficiary || undefined}><span className="table-cell-ellipsis">{row.beneficiary || "-"}</span></td>
                      <td title={row.employeeNumber || row.companyId || undefined}><span className="table-cell-ellipsis">{row.employeeNumber || row.companyId || "-"}</span></td>

                      {[0, 1, 2].map((index) => {
                        const milestone = row.milestones?.[index];

                        return (
                          <td key={`${row.screenCode}-${index}`} className="incentive-page__part-cell">
                            {milestone ? (
                              <button
                                type="button"
                                className={`incentive-page__part-button incentive-page__part-button--${partTone(
                                  milestone.status,
                                )}`}
                                onClick={() => openPayment(row, milestone)}
                                disabled={
                                  ![
                                    "Ready",
                                    "Downloaded",
                                    "Paid",
                                  ].includes(milestone.status)
                                }
                                title={partTooltip(milestone)}
                                aria-label={`${milestone.label}: ${money(
                                  milestone.amount,
                                )}. ${milestone.status}`}
                              >
                                {money(milestone.amount)}
                              </button>
                            ) : (
                              "-"
                            )}
                          </td>
                        );
                      })}

                      <td className="incentive-page__action-cell">
                        <button
                          type="button"
                          className="incentive-page__secondary-button incentive-page__action-button"
                          onClick={() => setActiveScreen(row.screenCode)}
                          title="View Details"
                          aria-label="View Details"
                        >
                          <svg
                            viewBox="0 0 24 24"
                            aria-hidden="true"
                            focusable="false"
                          >
                            <path d="M2.5 12s3.5-6 9.5-6 9.5 6 9.5 6-3.5 6-9.5 6-9.5-6-9.5-6Z" />
                            <circle cx="12" cy="12" r="2.5" />
                          </svg>
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="incentive-page__pagination">
              <span>
                Showing {firstVisible}-{lastVisible} of {filtered.length}
              </span>

              <div className="incentive-page__pagination-actions">
                <button
                  type="button"
                  className="incentive-page__secondary-button"
                  disabled={currentPage <= 1}
                  onClick={() =>
                    setCurrentPage((page) => Math.max(1, page - 1))
                  }
                >
                  Previous
                </button>

                <span>
                  Page {currentPage} of {totalPages}
                </span>

                <button
                  type="button"
                  className="incentive-page__secondary-button"
                  disabled={currentPage >= totalPages}
                  onClick={() =>
                    setCurrentPage((page) => Math.min(totalPages, page + 1))
                  }
                >
                  Next
                </button>
              </div>
            </div>
          </>
        )}
      </article>

      {activeRow ? (
        <article className="incentive-page__card incentive-page__card--details">
          <div className="incentive-page__card-header">
            <div>
              <p className="incentive-page__card-title">
                Site Incentive Details
              </p>
              <span>{activeRow.screenCode}</span>
            </div>

            <button
              type="button"
              className="incentive-page__secondary-button"
              onClick={() => setActiveScreen("")}
            >
              Close
            </button>
          </div>

          <div className="incentive-page__summary-grid">
            <div>
              <span>Screen Name</span>
              <strong>{activeRow.screenName || "-"}</strong>
            </div>
            <div>
              <span>Location</span>
              <strong>{activeRow.location || "-"}</strong>
            </div>
            <div>
              <span>State</span>
              <strong>{activeRow.state || "-"}</strong>
            </div>
            <div>
              <span>Beneficiary</span>
              <strong>{activeRow.beneficiary || "-"}</strong>
            </div>
            <div>
              <span>Company ID</span>
              <strong>
                {activeRow.employeeNumber || activeRow.companyId || "-"}
              </strong>
            </div>
            <div>
              <span>Billing Start Date</span>
              <strong>{displayDate(activeRow.billingStartDate)}</strong>
            </div>
          </div>

          <div className="incentive-page__milestone-list">
            {(activeRow.milestones || []).map((milestone) => (
              <section
                className="incentive-page__milestone-card"
                key={milestone.key}
              >
                <div className="incentive-page__milestone-heading">
                  <div>
                    <h3>{milestone.label}</h3>
                    <p>{money(milestone.amount)}</p>
                  </div>

                  <span
                    className={`incentive-page__status incentive-page__status--${milestone.status
                      .toLowerCase()
                      .replace(/\s+/g, "-")}`}
                  >
                    {milestone.status}
                  </span>
                </div>

                <div className="incentive-page__due-date">
                  <div>
                    <strong>Required Block:</strong> {milestone.blockLabel}
                  </div>
                  <div>
                    <strong>Payment Validation:</strong>{" "}
                    {milestone.paymentCount}/{milestone.requiredPaymentCount}{" "}
                    required payments received
                  </div>
                  <div>
                    <strong>Block Complete Date:</strong>{" "}
                    {displayDate(milestone.blockEndDate)}
                  </div>
                  <div>
                    <strong>Eligibility Date:</strong>{" "}
                    {milestone.eligibilityDate
                      ? displayDate(milestone.eligibilityDate)
                      : "-"}
                  </div>

                  {milestone.paidDate ? (
                    <div>
                      <strong>Incentive Payment Date:</strong>{" "}
                      {displayDate(milestone.paidDate)}
                    </div>
                  ) : null}

                  {milestone.paymentRemarks ? (
                    <div>
                      <strong>Payment Remarks:</strong>{" "}
                      {milestone.paymentRemarks}
                    </div>
                  ) : null}

                  {milestone.noBillingKeys?.length ? (
                    <div>
                      <strong>No Billing months excluded:</strong>{" "}
                      {milestone.noBillingKeys.join(", ")}
                    </div>
                  ) : null}
                </div>

                {milestone.missing?.length ? (
                  <div className="incentive-page__missing">
                    <strong>Pending requirements</strong>
                    <ul>
                      {milestone.missing.map((message) => (
                        <li key={message}>{message}</li>
                      ))}
                    </ul>
                  </div>
                ) : (
                  <p className="incentive-page__ready-message">
                    All required conditions are available.
                  </p>
                )}

                {milestone.status === "Review Required" ? (
                  <p className="incentive-page__review-message">
                    Critical source information changed after processing.
                    Manual review is required.
                  </p>
                ) : null}
              </section>
            ))}
          </div>
        </article>
      ) : null}

      {paymentTarget ? (
        <div
          className="incentive-page__modal-backdrop"
          role="presentation"
          onMouseDown={(event) => {
            if (event.target === event.currentTarget) closePayment();
          }}
        >
          <section
            className="incentive-page__payment-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="incentive-payment-title"
          >
            <div className="incentive-page__payment-modal-header">
              <div>
                <h2 id="incentive-payment-title">Incentive Payment</h2>
                <span>
                  {paymentTarget.row.screenCode} · {paymentTarget.milestone.label}
                </span>
              </div>
              <strong className="incentive-page__payment-amount">
                {money(paymentTarget.milestone.amount)}
              </strong>
            </div>

            <div className="incentive-page__payment-form">
              <div className="incentive-page__field">
                <label htmlFor="incentive-payment-date">Payment Date *</label>
                <input
                  id="incentive-payment-date"
                  type="date"
                  value={paymentDate}
                  min={
                    paymentTarget.milestone.eligibilityDate
                      ? n(paymentTarget.milestone.eligibilityDate).slice(0, 10)
                      : undefined
                  }
                  max={todayInputValue()}
                  onChange={(event) => {
                    setPaymentDate(event.target.value);
                    setApiError("");
                  }}
                  readOnly={paymentTarget.milestone.status === "Paid"}
                  disabled={paymentBusy}
                />
              </div>

              <div className="incentive-page__field incentive-page__field--remarks">
                <label htmlFor="incentive-payment-remarks">
                  Payment Remarks *
                </label>
                <textarea
                  id="incentive-payment-remarks"
                  value={paymentRemarks}
                  onChange={(event) => {
                    setPaymentRemarks(event.target.value);
                    setApiError("");
                  }}
                  placeholder="Enter brief payment details"
                  maxLength={PAYMENT_REMARKS_MAX_LENGTH}
                  aria-describedby="incentive-payment-remarks-helper"
                  readOnly={paymentTarget.milestone.status === "Paid"}
                  disabled={paymentBusy}
                />
                <span
                  id="incentive-payment-remarks-helper"
                  style={{
                    display: "block",
                    marginTop: 6,
                    color: "#9fb4c5",
                    fontSize: 12,
                    lineHeight: 1.4,
                  }}
                >
                  Payment reference, invoice/receipt details, or short note
                </span>
              </div>
            </div>

            {apiError ? (
              <div
                className="incentive-page__api-message incentive-page__api-message--error"
                role="alert"
                style={{ marginTop: 14 }}
              >
                <strong>Incentive:</strong> {apiError}
              </div>
            ) : null}

            <div className="incentive-page__payment-actions">
              {paymentTarget.milestone.status !== "Paid" ? (
                <button
                  type="button"
                  className="incentive-page__primary-button"
                  onClick={savePayment}
                  disabled={paymentBusy || !paymentDate}
                >
                  {paymentBusy ? "Saving..." : "Save"}
                </button>
              ) : null}

              <button
                type="button"
                className="incentive-page__secondary-button"
                onClick={closePayment}
                disabled={paymentBusy}
              >
                Close
              </button>
            </div>
          </section>
        </div>
      ) : null}
    </section>
  );
}

export default Incentive;
