import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

function n(value) {
  return String(value ?? "").trim();
}

function incentiveParticipates(record) {
  const applicability = n(record?.incentiveBeneficiaryApplicable).toLowerCase();
  if (["not applicable", "no", "false"].includes(applicability)) return false;

  const eligibleSiteIds = Array.isArray(record?.incentiveEligibleSiteIds)
    ? record.incentiveEligibleSiteIds.map((value) => n(value)).filter(Boolean)
    : [];

  if (eligibleSiteIds.length === 0) {
    return applicability !== "applicable";
  }

  return eligibleSiteIds.includes(n(record?.screenCode));
}

function isFoCRecord(record) {
  if (!record) return false;
  if (record.focOrigin === true || record.foCOrigin === true) return true;

  const directValues = [
    record?.foc,
    record?.foC,
    record?.focApplicable,
    record?.isFoc,
    record?.isFoC,
    record?.freeOfCharge,
    record?.focStatus,
  ]
    .map((value) => n(value).toLowerCase())
    .filter(Boolean);

  if (
    directValues.some((value) =>
      ["yes", "true", "1", "foc", "free of charge"].includes(value),
    )
  ) {
    return true;
  }

  return [
    record?.originalCommercialStatus,
    record?.commercialStatus,
    record?.billingMode,
    record?.pricingStatus,
  ]
    .map((value) => n(value).toLowerCase())
    .some((value) => value === "foc" || value === "free of charge");
}

function baseExportRow(record, stage) {
  return {
    Stage: stage,
    "Billing ID": n(record?.billingCode),
    "Complex ID": n(record?.complexCode) || "Standalone",
    "Billing Name": n(record?.billingName),
    "Screen Code": n(record?.screenCode),
    "Screen Name": n(record?.screenName),
    Location: n(record?.location),
    State: n(record?.state),
    "Site Type": n(record?.siteType),
    "Subscription Type": n(record?.subscriptionType),
    "Subscription Mode": n(record?.subscriptionMode),
    "Subscription Fee": n(record?.subscriptionFee),
    "OTF Applicable": n(record?.otfApplicable),
    "OTF Amount": n(record?.otfAmount),
    "Billing Start Date": n(record?.billingStartDate),
    "Current Status":
      n(record?.currentBillingStatus) ||
      n(record?.currentStageStatus) ||
      n(record?.installationStatus) ||
      n(record?.processingStatus),
  };
}

export default function Downloads({
  stage1Records = [],
  stage2Records = [],
  stage3Records = [],
  billingRecords = [],
  incentiveStates = {},
}) {
  const [scope, setScope] = useState("Overall");
  const [search, setSearch] = useState("");

  const stage1Rows = useMemo(
    () => stage1Records.map((record) => baseExportRow(record, "Site Entry")),
    [stage1Records],
  );

  const stage2Rows = useMemo(
    () =>
      stage2Records.map((record) => ({
        ...baseExportRow(record, "Installations"),
        "Installation Date": n(record?.installationDate),
        "Live Date": n(record?.liveDate),
        "Trial Period": n(record?.trialPeriod),
        "Trial Extension": n(
          record?.totalTrialPeriodExtension || record?.trialPeriodExtension,
        ),
      })),
    [stage2Records],
  );

  const stage3Rows = useMemo(
    () =>
      stage3Records.map((record) => ({
        ...baseExportRow(record, "Billings"),
        "Billing Period From": n(record?.billingPeriodFrom),
        "Billing Period To": n(record?.billingPeriodTo),
        "Invoice No": n(record?.invoiceNumber),
        "Invoice Date": n(record?.invoiceDate),
        "Payment Status": n(record?.paymentStatus),
        "Payment Received Date": n(record?.paymentReceivedDate),
      })),
    [stage3Records],
  );

  const incentiveRows = useMemo(
    () =>
      stage3Records
        .filter((record) => incentiveParticipates(record))
        .filter((record) => !isFoCRecord(record))
        .map((record) => {
          const state = incentiveStates[n(record?.screenCode)] || {};
          const downloadedParts =
            state?.downloadedParts && typeof state.downloadedParts === "object"
              ? Object.keys(state.downloadedParts)
              : [];

          return {
            ...baseExportRow(record, "Incentive"),
            Beneficiary:
              n(record?.incentiveBeneficiary) ||
              n(record?.employeeName) ||
              n(record?.salesEmployeeName),
            "Company ID": n(record?.companyId),
            "OTF Structure": n(record?.otfApplicable),
            "Downloaded Incentive Parts": downloadedParts.join(", "),
          };
        }),
    [stage3Records, incentiveStates],
  );

  const paymentRows = useMemo(() => {
    const rows = [];

    billingRecords.forEach((record) => {
      const entries = Array.isArray(record?.invoiceEntries)
        ? record.invoiceEntries
        : [];

      if (!entries.length) {
        rows.push({
          Stage: "Billing Payment",
          "Billing ID": n(record?.billingCode),
          "Complex ID": n(record?.complexCode) || "Standalone",
          "Screen Code": n(record?.screenCode || record?.siteScope),
          "Screen Name": n(record?.screenName),
          "Period From": n(record?.billingPeriodFrom),
          "Period To": n(record?.billingPeriodTo),
          "Invoice No": n(record?.invoiceNumber),
          "Invoice Date": n(record?.invoiceDate),
          "Payment Status": n(record?.paymentStatus),
          "Receipt No": n(
            record?.receiptNumber ||
              record?.paymentReceiptNumber ||
              record?.erpReceiptNumber,
          ),
          "Payment Received Date": n(
            record?.paymentReceivedDate ||
              record?.receiptDate ||
              record?.erpReceiptDate,
          ),
        });
        return;
      }

      entries.forEach((entry) => {
        rows.push({
          Stage: "Billing Payment",
          "Billing ID": n(record?.billingCode),
          "Complex ID": n(record?.complexCode) || "Standalone",
          "Screen Code": n(record?.screenCode || record?.siteScope),
          "Screen Name": n(record?.screenName),
          "Period From": n(entry?.periodFrom || record?.billingPeriodFrom),
          "Period To": n(entry?.periodTo || record?.billingPeriodTo),
          "Invoice No": n(entry?.invoiceNumber || record?.invoiceNumber),
          "Invoice Date": n(entry?.invoiceDate || record?.invoiceDate),
          "Payment Status": n(entry?.paymentStatus || record?.paymentStatus),
          "Receipt No": n(
            entry?.receiptNumber ||
              record?.receiptNumber ||
              record?.erpReceiptNumber,
          ),
          "Payment Received Date": n(
            entry?.paymentReceivedDate ||
              entry?.receiptDate ||
              record?.paymentReceivedDate ||
              record?.receiptDate,
          ),
        });
      });
    });

    return rows;
  }, [billingRecords]);

  const availableRows = useMemo(() => {
    switch (scope) {
      case "Site Entry":
        return stage1Rows;
      case "Installations":
        return stage2Rows;
      case "Billings":
        return [...stage3Rows, ...paymentRows];
      case "Incentive":
        return incentiveRows;
      default:
        return [
          ...stage1Rows,
          ...stage2Rows,
          ...stage3Rows,
          ...paymentRows,
          ...incentiveRows,
        ];
    }
  }, [
    scope,
    stage1Rows,
    stage2Rows,
    stage3Rows,
    paymentRows,
    incentiveRows,
  ]);

  const filteredRows = useMemo(() => {
    const q = n(search).toLowerCase();
    if (!q) return availableRows;

    return availableRows.filter((row) =>
      Object.values(row)
        .join(" ")
        .toLowerCase()
        .includes(q),
    );
  }, [availableRows, search]);

  function downloadExcel() {
    if (!filteredRows.length) {
      alert("No records are available for download.");
      return;
    }

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(filteredRows);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Billing and INC");

    const safeScope = scope.replace(/\s+/g, "_");
    const today = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(
      workbook,
      `Billing_and_INC_${safeScope}_${today}.xlsx`,
    );
  }

  const inputStyle = {
    minHeight: 40,
    border: "1px solid #d7dce5",
    borderRadius: 8,
    padding: "8px 10px",
    background: "#ffffff",
    color: "#111827",
  };

  return (
    <div
      style={{
        padding: 24,
        width: "100%",
        minHeight: "100%",
        boxSizing: "border-box",
        background: "#ffffff",
        color: "#111827",
      }}
    >
      <div style={{ marginBottom: 18 }}>
        <div
          style={{
            fontSize: 12,
            textTransform: "uppercase",
            letterSpacing: ".08em",
            color: "#6b7280",
            marginBottom: 6,
          }}
        >
          Governance
        </div>
        <h1 style={{ margin: 0, fontSize: 28, color: "#111827" }}>Downloads</h1>
        <p style={{ margin: "8px 0 0", color: "#667085" }}>
          Download stage-wise or overall Billing & INC records. Search can be
          used to limit the export to a Billing ID, Complex ID, screen, or
          location.
        </p>
      </div>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "220px minmax(300px, 1fr) 170px",
          gap: 10,
          marginBottom: 16,
        }}
      >
        <select
          style={inputStyle}
          value={scope}
          onChange={(event) => setScope(event.target.value)}
        >
          <option value="Overall">Overall</option>
          <option value="Site Entry">Site Entry</option>
          <option value="Installations">Installations</option>
          <option value="Billings">Billings</option>
          <option value="Incentive">Incentive</option>
        </select>

        <input
          style={inputStyle}
          type="search"
          placeholder="Filter by Billing ID, Complex ID, Screen Code, Screen Name, Location..."
          value={search}
          onChange={(event) => setSearch(event.target.value)}
        />

        <button
          type="button"
          onClick={downloadExcel}
          style={{
            border: 0,
            borderRadius: 8,
            background: "#222222",
            color: "#fff",
            fontWeight: 600,
            cursor: "pointer",
          }}
        >
          Download Excel
        </button>
      </div>

      <div
        style={{
          border: "1px solid #e1e5ec",
          borderRadius: 12,
          background: "#ffffff",
          padding: 18,
        }}
      >
        <div style={{ fontSize: 13, color: "#667085" }}>Selected scope</div>
        <div style={{ fontSize: 22, fontWeight: 700, marginTop: 4 }}>
          {scope}
        </div>
        <div style={{ marginTop: 12 }}>
          <strong>{filteredRows.length}</strong> record
          {filteredRows.length === 1 ? "" : "s"} will be exported.
        </div>
      </div>
    </div>
  );
}
