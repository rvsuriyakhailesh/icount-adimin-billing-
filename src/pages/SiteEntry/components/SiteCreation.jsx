import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import "./SiteCreation.css";
import {
  getIdValidationMessage,
  isStandaloneComplexCode,
  isValidIdInput,
  normalizeIdInput,
} from "../../../utils/idValidation";
import { isCurrentDownstreamWorkflowStage } from "../../../utils/workflowStage";

const GUIDED_SCROLL_DURATION_MS = 800;
const GUIDED_SCROLL_TOP_OFFSET_PX = 24;

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

function createEmptySiteRow() {
  return {
    id: crypto.randomUUID(),
    billingCode: "",
    complexCode: "",
    billingName: "",
    screenCode: "",
    screenName: "",
    location: "",
    state: "",
    siteType: "",
    statusType: "",
    statusText: "",
    previousStatus: "",
    currentStatus: "",
    selected: false,
    status: "Incomplete",
    foc: false,
    manualEntry: true,
  };
}

const supportedColumns = [
  "Billing Code / Customer Code",
  "Billing Code",
  "Billing ID",
  "Customer Code",
  "Customer ID",
  "Cusotmer Code",
  "Cusotmer ID",
  "Complex Code",
  "Complex ID",
  "Billing Name",
  "Complex Name",
  "Screen Code",
  "Screen ID",
  "Site ID",
  "Screen Name",
  "ERP Screen Name",
  "Location",
  "City / Town",
  "State",
  "Site Type",
  "Status Type",
  "Status Text",
  "Previous Status",
  "Current Status",
];

const billingCodeHeaderAliases = [
  "billing code",
  "billing id",
  "customer code",
  "customer id",
  "cusotmer code",
  "cusotmer id",
];
const complexCodeHeaderAliases = [
  "complex code",
  "complex id",
  "complexcode",
  "complexid",
];
const screenCodeHeaderAliases = [
  "screen code",
  "screen id",
  "site id",
  "screencode",
  "screenid",
];

const requiredColumnGroups = [
  { label: "Billing Code / Customer Code", aliases: billingCodeHeaderAliases },
  { label: "Screen Code / Screen ID", aliases: screenCodeHeaderAliases },
  { label: "Screen Name or ERP Screen Name", aliases: ["screen name", "erp screen name"] },
  { label: "Location or City / Town", aliases: ["location", "city / town"] },
  { label: "State", aliases: ["state"] },
];

const acceptedSiteTypes = ["Standalone", "Complex"];

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ");
}

function normalizeValue(value) {
  return String(value || "").trim();
}

function normalizeSiteType(value) {
  const normalizedValue = normalizeValue(value);
  if (!normalizedValue) {
    return "";
  }

  return normalizedValue.charAt(0).toUpperCase() + normalizedValue.slice(1).toLowerCase();
}

function inferSiteType(complexCode) {
  return isStandaloneComplexCode(complexCode) ? "Standalone" : "Complex";
}

function isRowEmpty(rowObject) {
  return Object.values(rowObject).every((value) => String(value || "").trim() === "");
}

function getUniqueNormalizedValues(rows, field, { allowBlank = false } = {}) {
  const values = rows
    .map((row) => normalizeValue(row?.[field]))
    .filter((value) => allowBlank || value);

  return Array.from(
    new Set(values.map((value) => value.toUpperCase())),
  );
}

function getNormalizedIdValue(value, { allowStandaloneBlank = false } = {}) {
  return normalizeIdInput(value, { allowStandaloneBlank });
}

function isBlankSiteRow(row) {
  return !normalizeIdInput(row?.screenCode) && !normalizeValue(row?.screenName);
}

function normalizeEligibilityValue(value) {
  return normalizeValue(value).toLowerCase();
}

function getSiteEligibility(row) {
  const statusType = normalizeEligibilityValue(
    row?.statusType || row?.eligibilityStatus || row?.statusCode,
  );
  const previousStatus = normalizeValue(row?.previousStatus);
  const currentStatus = normalizeValue(row?.currentStatus || row?.statusText || row?.status);
  const statusBlob = [
    statusType,
    previousStatus,
    currentStatus,
    normalizeValue(row?.siteType),
  ]
    .join(" ")
    .toLowerCase();

  if (!normalizeIdInput(row?.screenCode) && !normalizeValue(row?.screenName)) {
    return {
      selectable: false,
      badgeClass: "ignored",
      label: "Ignored - Empty Row",
      rowClass: "site-creation__row--disabled",
    };
  }

  if (statusType === "renewal" || statusBlob.includes("renewal")) {
    return {
      selectable: false,
      badgeClass: "renewal",
      label: "Renewal Workflow Only",
      rowClass: "site-creation__row--disabled",
    };
  }

  if (
    statusType === "rejoin" ||
    statusBlob.includes("contract quit") ||
    statusBlob.includes("inactive") ||
    statusBlob.includes("removed")
  ) {
    return {
      selectable: true,
      badgeClass: "rejoin",
      label: "Re-onboarding - Selectable",
      rowClass: "site-creation__row--rejoin",
    };
  }

  if (
    statusType === "active" ||
    statusBlob.includes("active") ||
    statusBlob.includes("downstream") ||
    statusBlob.includes("live") ||
    statusBlob.includes("billing")
  ) {
    return {
      selectable: false,
      badgeClass: "active",
      label: "Existing - Disabled",
      rowClass: "site-creation__row--disabled",
    };
  }

  return {
    selectable: true,
    badgeClass: "new",
    label: "New or Renewal",
    rowClass: "",
  };
}

function getSiteStateBadge(row) {
  const statusType = normalizeEligibilityValue(
    row?.statusType || row?.eligibilityStatus || row?.statusCode,
  );
  const statusBlob = [
    statusType,
    normalizeValue(row?.previousStatus),
    normalizeValue(row?.currentStatus || row?.statusText || row?.status),
  ]
    .join(" ")
    .toLowerCase();

  if (
    statusType === "renewal" ||
    statusType === "rejoin" ||
    statusBlob.includes("renewal") ||
    statusBlob.includes("contract quit") ||
    statusBlob.includes("inactive") ||
    statusBlob.includes("removed")
  ) {
    return {
      className: "state-renewal",
      label: "Renewal",
    };
  }

  return {
    className: "state-new",
    label: "New",
  };
}


function SiteCreation({
  onDataChange = () => {},
  resetSignal = 0,
  editGroupId = "",
  editGroupRecords = [],
  blockedSiteIds = [],
  backendSiteRecords = [],
  isFullscreenActive = false,
  onRequestFullscreenRestore = () => {},
  onImportCompleted = () => {},
}) {
  const [billingCode, setBillingId] = useState("");
  const [complexCode, setComplexId] = useState("");
  const [billingName, setBillingName] = useState("");
  const [location, setLocation] = useState("");
  const [state, setState] = useState("");
  const [processingStatus, setProcessingStatus] = useState("New");
  const [siteRows, setSiteRows] = useState([createEmptySiteRow()]);
  const [uploadPreview, setUploadPreview] = useState(null);
  const [uploadError, setUploadError] = useState("");
  const fileInputRef = useRef(null);
  const uploadPreviewRef = useRef(null);
  const currentBillingIdRef = useRef("");
  const fullscreenWasActiveBeforePickerRef = useRef(false);
  const blockedSiteIdSet = useMemo(
    () =>
      new Set(
        blockedSiteIds
          .map((screenCode) => normalizeIdInput(screenCode))
          .filter(Boolean),
      ),
    [blockedSiteIds],
  );
  const activeDownstreamBackendSiteIdSet = useMemo(
    () =>
      new Set(
        backendSiteRecords
          .filter((site) => isCurrentDownstreamWorkflowStage(site))
          .map((site) => normalizeIdInput(site?.siteId))
          .filter(Boolean),
      ),
    [backendSiteRecords],
  );

  function getEffectiveSiteEligibility(row) {
    const screenCode = normalizeIdInput(row?.screenCode);

    if (screenCode && activeDownstreamBackendSiteIdSet.has(screenCode)) {
      return {
        selectable: false,
        badgeClass: "active",
        label: "Existing",
        rowClass: "site-creation__row--disabled",
      };
    }

    if (screenCode && blockedSiteIdSet.has(screenCode)) {
      return {
        selectable: false,
        badgeClass: "active",
        label: "Existing",
        rowClass: "site-creation__row--disabled",
      };
    }

    return getSiteEligibility(row);
  }

  const selectableRows = siteRows.filter(
    (row) => getEffectiveSiteEligibility(row).selectable,
  );
  const hasSelectableSiteRow = selectableRows.length > 0;

  useEffect(() => {
    onDataChange({
      billingCode,
      complexCode,
      billingName,
      location,
      state,
      processingStatus,
      siteRows,
    });
  }, [
    billingCode,
    complexCode,
    billingName,
    location,
    state,
    processingStatus,
    siteRows,
    onDataChange,
  ]);

  useEffect(() => {
    if (resetSignal === 0) {
      return;
    }

    setBillingId("");
    setComplexId("");
    setBillingName("");
    setLocation("");
    setState("");
    setProcessingStatus("New");
    setSiteRows([createEmptySiteRow()]);
    setUploadPreview(null);
    setUploadError("");
    currentBillingIdRef.current = "";

    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  }, [resetSignal]);

  useEffect(() => {
    if (!editGroupId) {
      return;
    }

    const sourceGroupRecord = editGroupRecords[0];
    const snapshotRows = sourceGroupRecord?.stage1GroupRows;
    const nextRows =
      Array.isArray(snapshotRows) && snapshotRows.length > 0
        ? snapshotRows.map((row) => ({
            ...createEmptySiteRow(),
            id: crypto.randomUUID(),
            billingCode: row.billingCode || "",
            complexCode: row.complexCode || "",
            billingName: row.billingName || sourceGroupRecord?.billingName || "",
            screenCode: row.screenCode || "",
            screenName: row.screenName || "",
            location: row.location || "",
            state: row.state || "",
            siteType:
              row.siteType ||
              (normalizeValue(row.complexCode) ? "Complex" : "Standalone"),
            statusType: row.statusType || "",
            statusText: row.statusText || "",
            previousStatus: row.previousStatus || "",
            currentStatus: row.currentStatus || "",
            selected: false,
            status: row.status || "Complete",
            stage1CorrectionReturn: Boolean(
              row.stage1CorrectionReturn || sourceGroupRecord?.stage1CorrectionReturn,
            ),
            stage1PreviouslySaved: true,
            foc: Boolean(row.foc),
            manualEntry: false,
          }))
        : editGroupRecords.map((record) => ({
            ...createEmptySiteRow(),
            id: crypto.randomUUID(),
            billingCode: record.billingCode || "",
            complexCode: record.complexCode || "",
            billingName: record.billingName || "",
            screenCode: record.screenCode || "",
            screenName: record.screenName || "",
            location: record.location || "",
            state: record.state || "",
            siteType:
              record.siteType ||
              (normalizeValue(record.complexCode) ? "Complex" : "Standalone"),
            statusType: record.statusType || "",
            statusText: record.statusText || "",
            previousStatus: record.previousStatus || "",
            currentStatus: record.currentStatus || "",
            selected: false,
            status: record.status || "Complete",
            stage1CorrectionReturn: Boolean(record.stage1CorrectionReturn),
            stage1PreviouslySaved: true,
            foc: Boolean(record.foc),
            manualEntry: false,
          }));

    const firstRow = nextRows[0] || createEmptySiteRow();

    setBillingId(normalizeIdInput(firstRow.billingCode));
    setComplexId(normalizeIdInput(firstRow.complexCode, { allowStandaloneBlank: true }));
    setBillingName(firstRow.billingName || sourceGroupRecord?.billingName || "");
    setLocation(firstRow.location || "");
    setState(firstRow.state || "");
    setProcessingStatus("Existing");
    setSiteRows(nextRows.length > 0 ? nextRows : [createEmptySiteRow()]);
    setUploadPreview(null);
    setUploadError("");
    currentBillingIdRef.current = normalizeIdInput(firstRow.billingCode);

    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  }, [editGroupId, editGroupRecords]);

  function getHeaderMap(headerRow) {
    return headerRow.reduce((map, header, index) => {
      const normalizedHeader = normalizeHeader(header);
      if (!normalizedHeader) {
        return map;
      }

      map[normalizedHeader] = [...(map[normalizedHeader] || []), index];
      return map;
    }, {});
  }

  function getColumnIndices(headerMap, aliases) {
    return Array.from(
      new Set(
        aliases.flatMap((alias) =>
          headerMap[normalizeHeader(alias)] || [],
        ),
      ),
    );
  }

  function getColumnIndex(headerMap, aliases) {
    return getColumnIndices(headerMap, aliases)[0];
  }

  function getColumnValue(row, headerMap, aliases) {
    const index = getColumnIndex(headerMap, aliases);
    return index === undefined ? "" : normalizeValue(row[index]);
  }

  function getAliasedIdValue(row, headerMap, aliases, fieldLabel, options) {
    const values = Array.from(
      new Set(
        getColumnIndices(headerMap, aliases)
          .map((index) => normalizeIdInput(row[index], options))
          .filter(Boolean),
      ),
    );

    return {
      value: values[0] || "",
      error:
        values.length > 1
          ? `Conflicting values found for ${fieldLabel}.`
          : "",
    };
  }

  function buildRowObject(row, headerMap) {
    const billingCode = getAliasedIdValue(
      row,
      headerMap,
      billingCodeHeaderAliases,
      "Billing Code / Billing ID / Customer Code / Customer ID",
    );
    const complexCode = getAliasedIdValue(
      row,
      headerMap,
      complexCodeHeaderAliases,
      "Complex Code / Complex ID",
      { allowStandaloneBlank: true },
    );
    const screenCode = getAliasedIdValue(
      row,
      headerMap,
      screenCodeHeaderAliases,
      "Screen Code / Screen ID",
    );

    return {
      billingCode: billingCode.value,
      complexCode: complexCode.value,
      billingName: normalizeValue(
        getColumnValue(row, headerMap, ["billing name", "complex name"]),
      ),
      screenCode: screenCode.value,
      screenName: normalizeValue(getColumnValue(row, headerMap, ["screen name", "erp screen name"])),
      location: normalizeValue(getColumnValue(row, headerMap, ["location", "city / town"])),
      state: normalizeValue(getColumnValue(row, headerMap, ["state"])),
      siteType: getColumnValue(row, headerMap, ["site type"]),
      statusType: getColumnValue(row, headerMap, ["status type"]),
      statusText: getColumnValue(row, headerMap, ["status text"]),
      previousStatus: getColumnValue(row, headerMap, ["previous status"]),
      currentStatus: getColumnValue(row, headerMap, ["current status"]),
      headerErrors: [billingCode.error, complexCode.error, screenCode.error].filter(
        Boolean,
      ),
    };
  }

  function validateUploadRows(rows, headerMap, existingSiteIds) {
    const seenSiteIds = new Set();
    const validRows = [];
    const invalidRows = [];

    rows.forEach((row, rowIndex) => {
      const rowNumber = rowIndex + 2;
      const { headerErrors = [], ...rowData } = buildRowObject(row, headerMap);

      if (isRowEmpty(rowData)) {
        return;
      }

      const errors = [...headerErrors];
      const billingIdValue = normalizeIdInput(rowData.billingCode);
      const siteIdValue = normalizeIdInput(rowData.screenCode);
      const complexIdValue = normalizeIdInput(rowData.complexCode, {
        allowStandaloneBlank: true,
      });

      if (!isValidIdInput(billingIdValue)) {
        errors.push(getIdValidationMessage("Billing Code / Customer Code"));
      }

      if (!isValidIdInput(siteIdValue)) {
        errors.push(getIdValidationMessage("Screen Code"));
      }

      if (!rowData.screenName) errors.push("Screen Name is mandatory.");
      if (!rowData.location) errors.push("Location is mandatory.");
      if (!rowData.state) errors.push("State is mandatory.");

      const normalizedComplexId = complexIdValue;
      const inferredSiteType = inferSiteType(normalizedComplexId);
      const resolvedSiteType = normalizeSiteType(rowData.siteType) || inferredSiteType;

      if (!acceptedSiteTypes.includes(resolvedSiteType)) {
        errors.push(
          `Site Type must be ${acceptedSiteTypes.join(", ")}.`,
        );
      } else {
        rowData.siteType = resolvedSiteType;
      }

      const normalizedSiteId = rowData.screenCode.toUpperCase();
      if (seenSiteIds.has(normalizedSiteId)) {
        errors.push("Duplicate Screen Code found in upload.");
      }
      if (existingSiteIds.has(normalizedSiteId)) {
        errors.push("Duplicate Screen Code already exists in the current Site Creation table.");
      }
      seenSiteIds.add(normalizedSiteId);

      if (errors.length > 0) {
        invalidRows.push({
          rowNumber,
          screenCode: rowData.screenCode,
          reason: errors.join(" "),
        });
      } else {
        validRows.push({
          rowNumber,
          ...rowData,
          billingCode: billingIdValue,
          complexCode: normalizedComplexId,
          screenCode: siteIdValue,
          screenName: rowData.screenName,
          location: rowData.location,
          state: rowData.state,
          siteType: resolvedSiteType,
        });
      }
    });

    return {
      validRows,
      invalidRows,
      summary: {
        totalRows: validRows.length + invalidRows.length,
        validRows: validRows.length,
        invalidRows: invalidRows.length,
      },
    };
  }

  function handleFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    if (
      fullscreenWasActiveBeforePickerRef.current &&
      document.fullscreenElement !== document.documentElement
    ) {
      fullscreenWasActiveBeforePickerRef.current = false;
      onRequestFullscreenRestore?.();
    } else {
      fullscreenWasActiveBeforePickerRef.current = false;
    }

    const fileName = file.name.toLowerCase();
    const isCsv = fileName.endsWith(".csv");
    const isXlsx = fileName.endsWith(".xlsx");

    if (!isCsv && !isXlsx) {
      setUploadError("Unsupported file format. Please upload .csv or .xlsx.");
      event.target.value = "";
      return;
    }

    const reader = new FileReader();

    reader.onload = (loadEvent) => {
      try {
        const fileData = loadEvent.target.result;
        const workbook = XLSX.read(fileData, {
          type: isCsv ? "string" : "array",
          raw: true,
        });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: "",
        });

        if (!rows || rows.length === 0) {
          setUploadError("The uploaded file is empty.");
          setUploadPreview(null);
          return;
        }

        const headerRow = rows[0].map((cell) => normalizeValue(cell));
        const headerMap = getHeaderMap(headerRow);
        const missingColumns = requiredColumnGroups.filter(
          (group) => getColumnIndex(headerMap, group.aliases) === undefined,
        );

        if (missingColumns.length > 0) {
          setUploadError(
            `Missing required columns: ${missingColumns.map((group) => group.label).join(", ")}.`,
          );
          setUploadPreview(null);
          return;
        }

        const dataRows = rows.slice(1);

        const billingNameColumnIndex = getColumnIndex(headerMap, ["billing name"]);
        const complexNameColumnIndex = getColumnIndex(headerMap, ["complex name"]);

        if (
          billingNameColumnIndex !== undefined &&
          complexNameColumnIndex !== undefined
        ) {
          const mismatchedNameRowIndex = dataRows.findIndex((row) => {
            const billingNameValue = normalizeValue(
              row[billingNameColumnIndex],
            ).toLowerCase();
            const complexNameValue = normalizeValue(
              row[complexNameColumnIndex],
            ).toLowerCase();

            return (
              billingNameValue &&
              complexNameValue &&
              billingNameValue !== complexNameValue
            );
          });

          if (mismatchedNameRowIndex >= 0) {
            setUploadError(
              `Billing Name and Complex Name do not match in Excel row ${
                mismatchedNameRowIndex + 2
              }. Please correct the values before upload.`,
            );
            setUploadPreview(null);
            return;
          }
        }

        const existingSiteIds = new Set(
          siteRows
            .map((row) => row.screenCode.trim().toUpperCase())
            .filter(Boolean),
        );
        const { validRows, invalidRows, summary } = validateUploadRows(
          dataRows,
          headerMap,
          existingSiteIds,
        );

        const incomingBillingIds = getUniqueNormalizedValues(validRows, "billingCode");
        const incomingBillingId =
          incomingBillingIds.length === 1 ? incomingBillingIds[0] : "";

        if (incomingBillingId && shouldConfirmBillingReplacement(incomingBillingId)) {
          const currentBillingId = normalizeIdInput(currentBillingIdRef.current);
          const confirmed = window.confirm(
            `1A already contains site information for Billing Code / Customer Code ${currentBillingId}. Using ${incomingBillingId} will replace the current 1A site information. Do you want to continue?`,
          );

          if (!confirmed) {
            return;
          }

          replaceActiveBillingContext(incomingBillingId);
        }

        setUploadPreview({
          fileName: file.name,
          billingCode: incomingBillingId,
          validRows,
          invalidRows,
          summary,
        });

        window.requestAnimationFrame(() => {
          guidedScrollToElement(uploadPreviewRef.current);
        });
      } catch (error) {
        setUploadError("Unable to parse the uploaded file. Please check the format.");
        setUploadPreview(null);
      } finally {
        event.target.value = "";
      }
    };

    if (isCsv) {
      reader.readAsText(file);
    } else {
      reader.readAsArrayBuffer(file);
    }
  }

  function handleCancelUpload() {
    setUploadPreview(null);
    setUploadError("");
  }

  function handleImportValidRows() {
    if (!uploadPreview?.validRows?.length) {
      return;
    }

    const existingSiteIds = new Set(
      siteRows.map((row) => row.screenCode.trim().toUpperCase()).filter(Boolean),
    );

    const rowsToImport = uploadPreview.validRows.filter(
      (row) => !existingSiteIds.has(row.screenCode.toUpperCase()),
    );

    if (rowsToImport.length === 0) {
      alert("No valid rows are available for import after removing duplicates.");
      return;
    }

    const billingIds = getUniqueNormalizedValues(rowsToImport, "billingCode");
    if (billingIds.length !== 1) {
      alert("All imported rows must use the same Billing Code / Customer Code.");
      return;
    }

    const locations = getUniqueNormalizedValues(rowsToImport, "location");
    if (locations.length !== 1) {
      alert("All imported rows must use the same Location.");
      return;
    }

    const states = getUniqueNormalizedValues(rowsToImport, "state");
    if (states.length !== 1) {
      alert("All imported rows must use the same State.");
      return;
    }

    const billingIdValue = normalizeValue(rowsToImport[0].billingCode);
    const locationValue = normalizeValue(rowsToImport[0].location);
    const stateValue = normalizeValue(rowsToImport[0].state);

    const importedComplexIds = Array.from(
      new Set(
        rowsToImport
          .map((row) =>
            normalizeIdInput(row.complexCode, {
              allowStandaloneBlank: true,
            }),
          )
          .filter(Boolean),
      ),
    );

    if (importedComplexIds.length > 1) {
      alert(
        "Imported rows contain more than one Complex Code. Please verify the file before importing.",
      );
      return;
    }

    const complexIdValue =
      importedComplexIds.length === 1 ? importedComplexIds[0] : "";

    const importedBillingNames = Array.from(
      new Set(
        rowsToImport
          .map((row) => normalizeValue(row.billingName))
          .filter(Boolean),
      ),
    );

    if (importedBillingNames.length > 1) {
      alert(
        "Imported rows contain more than one Billing Name. Please verify the file before importing.",
      );
      return;
    }

    const billingNameValue =
      importedBillingNames.length === 1 ? importedBillingNames[0] : "";

    setSiteRows((currentRows) => {
      const importedRows = rowsToImport.map((row) => ({
        ...createEmptySiteRow(),
        id: crypto.randomUUID(),
        screenCode: normalizeIdInput(row.screenCode),
        screenName: normalizeValue(row.screenName),
        billingCode: normalizeIdInput(row.billingCode),
        complexCode: normalizeIdInput(row.complexCode, { allowStandaloneBlank: true }),
        billingName: normalizeValue(row.billingName) || billingNameValue,
        location: normalizeValue(row.location),
        state: normalizeValue(row.state),
        siteType: row.siteType,
        statusType: row.statusType,
        statusText: row.statusText,
        previousStatus: row.previousStatus,
        currentStatus: row.currentStatus,
        manualEntry: false,
      }));

      const populatedRows = currentRows.filter((row) => !isBlankSiteRow(row));
      const blankRows = currentRows.filter((row) => isBlankSiteRow(row));

      return [...populatedRows, ...importedRows, ...blankRows];
    });

    setBillingId(normalizeIdInput(billingIdValue));
    setComplexId(complexIdValue);
    setBillingName(billingNameValue);
    setLocation(locationValue);
    setState(stateValue);
    currentBillingIdRef.current = normalizeIdInput(billingIdValue);

    setUploadPreview(null);
    setUploadError("");

    window.requestAnimationFrame(() => {
      onImportCompleted();
    });
  }

  function handleSiteRowChange(rowId, field, value) {
    if (field === "foc" && value === true) {
      const targetRow = siteRows.find((row) => row.id === rowId);

      if (
        !targetRow ||
        !normalizeValue(targetRow.screenCode) ||
        !normalizeValue(targetRow.screenName)
      ) {
        return;
      }
    }

    setSiteRows((currentRows) =>
      currentRows.map((row) =>
        row.id === rowId
          ? {
              ...row,
              [field]:
                field === "screenCode"
                  ? normalizeIdInput(value)
                  : value,
            }
          : row,
      ),
    );
  }

  function handleBillingNameChange(value) {
    setBillingName(value);
    setSiteRows((currentRows) =>
      currentRows.map((row) => ({
        ...row,
        billingName: value,
      })),
    );
  }

  function addSiteRow() {
    setSiteRows((currentRows) => [
      ...currentRows,
      {
        ...createEmptySiteRow(),
        id: crypto.randomUUID(),
        billingName,
      },
    ]);
  }

  function removeSiteRow(rowId) {
    setSiteRows((currentRows) => {
      const nextRows = currentRows.filter((row) => row.id !== rowId);
      return nextRows.length > 0 ? nextRows : [createEmptySiteRow()];
    });
  }

  function selectAllSites() {
    const selectableIds = new Set(selectableRows.map((row) => row.id));

    setSiteRows((currentRows) =>
      currentRows.map((row) =>
        selectableIds.has(row.id)
          ? { ...row, selected: true }
          : { ...row, selected: false },
      ),
    );
  }

  function clearSiteSelection() {
    setSiteRows((currentRows) =>
      currentRows.map((row) => ({ ...row, selected: false })),
    );
  }

  function replaceActiveBillingContext(nextBillingId) {
    setBillingId(nextBillingId);
    setComplexId("");
    setBillingName("");
    setLocation("");
    setState("");
    setProcessingStatus("New");
    setSiteRows([createEmptySiteRow()]);
    setUploadPreview(null);
    setUploadError("");
    currentBillingIdRef.current = nextBillingId;

    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  }

  function shouldConfirmBillingReplacement(nextBillingId) {
    const currentBillingId = normalizeIdInput(currentBillingIdRef.current);
    return Boolean(currentBillingId) && hasSelectableSiteRow && currentBillingId !== nextBillingId;
  }

  function handleFetchSiteInformation() {
    const nextBillingId = normalizeIdInput(billingCode);

    if (!nextBillingId) {
      alert(getIdValidationMessage("Billing Code / Customer Code"));
      return;
    }

    const currentBillingId = normalizeIdInput(currentBillingIdRef.current);
    const hasCurrentSiteContext = Boolean(currentBillingId) && hasSelectableSiteRow;

    if (hasCurrentSiteContext && currentBillingId !== nextBillingId) {
      const confirmed = window.confirm(
        `1A already contains site information for Billing Code / Customer Code ${currentBillingId}. Fetching ${nextBillingId} will replace the current 1A site information. Do you want to continue?`,
      );

      if (!confirmed) {
        setBillingId(currentBillingId);
        return;
      }

      replaceActiveBillingContext(nextBillingId);
    } else if (!hasCurrentSiteContext) {
      currentBillingIdRef.current = nextBillingId;
    }

    alert("ERP site-information fetching will be connected later.");
  }

  function handleUploadVerify() {
    setUploadError("");
    fullscreenWasActiveBeforePickerRef.current =
      isFullscreenActive ||
      document.fullscreenElement === document.documentElement;

    const fileInput = fileInputRef.current;

    if (!fileInput) {
      return;
    }

    if (typeof fileInput.showPicker === "function") {
      fileInput.showPicker();
      return;
    }

    fileInput.click();
  }


  return (
    <section className="site-creation">
      <input
        ref={fileInputRef}
        type="file"
        accept=".csv,.xlsx"
        onChange={handleFileChange}
        aria-hidden="true"
        tabIndex={-1}
        style={{
          position: "fixed",
          width: "1px",
          height: "1px",
          opacity: 0,
          pointerEvents: "none",
          left: "-9999px",
        }}
      />

      <div className="site-creation__toolbar">
        <button
          type="button"
          className="site-creation__button"
          onClick={handleFetchSiteInformation}
        >
          ERP Fetch
        </button>

        <button
          type="button"
          className="site-creation__button site-creation__button--secondary"
          onClick={handleUploadVerify}
        >
          Upload & Verify
        </button>
      </div>

      <div className="site-creation__form-grid">
        <label className="site-creation__field">
          <span>Billing Code / Customer Code</span>
          <input
            type="text"
            maxLength={8}
            value={billingCode}
            onChange={(event) => setBillingId(normalizeIdInput(event.target.value))}
            placeholder="Enter Billing Code / Customer Code"
          />
        </label>

        <label className="site-creation__field">
          <span>Complex Code</span>
          <input
            type="text"
            maxLength={8}
            value={complexCode}
            onChange={(event) =>
              setComplexId(
                normalizeIdInput(event.target.value, {
                  allowStandaloneBlank: true,
                }),
              )
            }
            placeholder="Optional for standalone site"
          />
        </label>

        <label className="site-creation__field">
          <span>Billing Name</span>
          <input
            type="text"
            value={billingName}
            onChange={(event) => handleBillingNameChange(event.target.value)}
            placeholder="Enter billing name"
          />
        </label>

        <label className="site-creation__field">
          <span>Location</span>
          <input
            type="text"
            value={location}
            onChange={(event) => setLocation(event.target.value)}
            placeholder="Enter location"
          />
        </label>

        <label className="site-creation__field">
          <span>State</span>
          <input
            type="text"
            value={state}
            onChange={(event) => setState(event.target.value)}
            placeholder="Enter state"
          />
        </label>

        <label className="site-creation__field">
          <span>Processing Status</span>
          <input type="text" value={processingStatus} readOnly />
        </label>
      </div>

      <div className="site-creation__table-toolbar">
        <div className="site-creation__table-actions">
          <button
            type="button"
            className="site-creation__text-button"
            onClick={selectAllSites}
          >
            Select All
          </button>

          <button
            type="button"
            className="site-creation__text-button"
            onClick={clearSiteSelection}
          >
            Clear Selection
          </button>
        </div>

      </div>

      {uploadError && (
        <div className="site-creation__upload-error">
          {uploadError}
        </div>
      )}

      {uploadPreview && (
        <div
          className="site-creation__upload-preview"
          ref={uploadPreviewRef}
        >
          <div className="site-creation__upload-header">
            <div>
              <p className="site-creation__upload-title">Upload Verification</p>
              <p className="site-creation__upload-subtitle">
                {uploadPreview.fileName}
              </p>
            </div>

            <div className="site-creation__upload-actions">
              <button
                type="button"
                className="site-creation__text-button"
                onClick={handleCancelUpload}
              >
                Cancel
              </button>
              <button
                type="button"
                className="site-creation__button"
                onClick={handleImportValidRows}
              >
                Import Valid Rows
              </button>
            </div>
          </div>

          <div className="site-creation__upload-summary">
            <div>
              <strong>Total rows</strong>
              <span>{uploadPreview.summary.totalRows}</span>
            </div>
            <div>
              <strong>Valid rows</strong>
              <span>{uploadPreview.summary.validRows}</span>
            </div>
            <div>
              <strong>Invalid rows</strong>
              <span>{uploadPreview.summary.invalidRows}</span>
            </div>
          </div>

          <div className="site-creation__upload-table-wrapper">
            <table className="site-creation__upload-table">
              <thead>
                <tr>
                  <th>Row</th>
                  <th>Screen Code</th>
                  <th>Error Reason</th>
                </tr>
              </thead>
              <tbody>
                {uploadPreview.invalidRows.length === 0 ? (
                  <tr>
                    <td colSpan="3">
                      No invalid rows found. You can import valid rows.
                    </td>
                  </tr>
                ) : (
                  uploadPreview.invalidRows.map((row) => (
                    <tr key={`${row.rowNumber}-${row.screenCode}`}>
                      <td>{row.rowNumber}</td>
                      <td>{row.screenCode || "-"}</td>
                      <td>{row.reason}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div className="site-creation__table-wrapper">
        <table className="site-creation__table">
          <colgroup>
            <col className="site-creation__col-select" />
            <col className="site-creation__col-screen-code" />
            <col className="site-creation__col-complex-code" />
            <col className="site-creation__col-site-name" />
            <col className="site-creation__col-foc" />
            <col className="site-creation__col-status" />
            <col className="site-creation__col-action" />
          </colgroup>
          <thead>
            <tr>
              <th>Select</th>
              <th>Screen Code</th>
              <th>Complex Code</th>
              <th>Site Name</th>
              <th>FoC</th>
              <th>Status</th>
              <th>Action</th>
            </tr>
          </thead>

          <tbody>
            {siteRows.map((row) => {
              const effectiveEligibility = getEffectiveSiteEligibility(row);
              const stateBadge =
                row.stage1PreviouslySaved === true ||
                row.stage1CorrectionReturn === true ||
                blockedSiteIdSet.has(normalizeIdInput(row.screenCode))
                  ? { className: "active", label: "Existing" }
                  : getSiteStateBadge(row);
              const isManualEntry = row.manualEntry === true;
              const isSelectable = effectiveEligibility.selectable;

              return (
                <tr key={row.id} className={effectiveEligibility.rowClass}>
                  <td>
                    <input
                      className="row-check"
                      type="checkbox"
                      checked={Boolean(row.selected)}
                      disabled={!isSelectable}
                      onChange={(event) =>
                        handleSiteRowChange(row.id, "selected", event.target.checked)
                      }
                      aria-label={`Select row ${row.screenCode || row.screenName || row.id}`}
                    />
                  </td>

                  <td>
                    {isManualEntry ? (
                      <input
                        className="site-id-input"
                        type="text"
                        maxLength={8}
                        value={row.screenCode}
                        onChange={(event) =>
                          handleSiteRowChange(
                            row.id,
                            "screenCode",
                            event.target.value,
                          )
                        }
                        placeholder="Screen Code"
                      />
                    ) : (
                      <span className="display-value table-cell-ellipsis" title={row.screenCode || undefined}>
                        {row.screenCode || "-"}
                      </span>
                    )}
                  </td>

                  <td>
                    <span className="display-value table-cell-ellipsis" title={normalizeValue(row.complexCode) || undefined}>
                      {normalizeValue(row.complexCode) || "-"}
                    </span>
                  </td>

                  <td>
                    {isManualEntry ? (
                      <input
                        className="site-name-input"
                        type="text"
                        value={row.screenName}
                        onChange={(event) =>
                          handleSiteRowChange(
                            row.id,
                            "screenName",
                            event.target.value,
                          )
                        }
                        placeholder="Site name"
                      />
                    ) : (
                      <span className="display-value table-cell-ellipsis" title={normalizeValue(row.screenName) || undefined}>
                        {normalizeValue(row.screenName) || "-"}
                      </span>
                    )}
                  </td>

                  <td>
                    <button
                      type="button"
                      className={`site-creation__foc-button ${
                        row.foc ? "site-creation__foc-button--active" : ""
                      }`}
                      onClick={() =>
                        handleSiteRowChange(row.id, "foc", !Boolean(row.foc))
                      }
                      disabled={
                        !normalizeValue(row.screenCode) ||
                        !normalizeValue(row.screenName)
                      }
                      aria-pressed={Boolean(row.foc)}
                      title={
                        !normalizeValue(row.screenCode) ||
                        !normalizeValue(row.screenName)
                          ? "FoC can be selected only after Screen Code and Site Name are available"
                          : row.foc
                            ? "FoC active - click to set No"
                            : "FoC not active - click to set Yes"
                      }
                    >
                      {row.foc ? "Yes" : "No"}
                    </button>
                  </td>

                  <td>
                    <span
                      className={`site-creation__state-badge site-creation__state-badge--${stateBadge.className}`}
                    >
                      {stateBadge.label}
                    </span>
                  </td>

                  <td>
                    <button
                      type="button"
                      className="remove-btn"
                      onClick={() => removeSiteRow(row.id)}
                    >
                      Remove
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>



    </section>
  );
}

export default SiteCreation;
