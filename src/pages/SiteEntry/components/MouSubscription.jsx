import { useEffect, useMemo, useRef, useState } from "react";
import "./MouSubscription.css";
import { sanitizeAmount } from "../../../utils/numberInput";
import { isValidFourDigitYear, isValidDateInputValue, isValidDateValue } from "../../../utils/dateValidation";

function normalizeValue(value) {
  return String(value || "").trim();
}

function getAllocationChainName(model, fallback = "") {
  return (
    normalizeValue(model?.chainName) ||
    normalizeValue(model?.networkName) ||
    normalizeValue(model?.name) ||
    normalizeValue(model?.modelName) ||
    normalizeValue(fallback)
  );
}

function isFoCSite(row) {
  const value =
    row?.foc ??
    row?.foC ??
    row?.focApplicable ??
    row?.isFoc ??
    row?.isFoC ??
    row?.freeOfCharge;

  if (typeof value === "boolean") return value;

  return ["yes", "true", "1", "foc", "free of charge"].includes(
    normalizeValue(value).toLowerCase(),
  );
}

const PRICING_METHOD_ALLOCATION = "Allocation Model";
const ALLOCATION_SETTINGS_KEY = "billing_inc_allocation_model_config";
const ALLOCATION_MODEL_VERSIONS_KEY = "billing_inc_allocation_model_versions";
const DEFAULT_ALLOCATION_SETTINGS = {
  modelId: "",
  chainName: "",
  modelName: "",
  status: "",
  effectiveFrom: "",
  otfPricePerDevice: "",
  upTo3PlanFee: "",
  moreThan3PlanFee: "",
  slabs: [],
};

function normalizeAllocationSettingsModel(model, fallbackName = "Allocation Model") {
  if (!model || typeof model !== "object") return null;

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
    ...DEFAULT_ALLOCATION_SETTINGS,
    ...model,
    modelId:
      normalizeValue(model.modelId) ||
      normalizeValue(model.id) ||
      `allocation-model-${normalizeValue(model.effectiveFrom) || "legacy"}`,
    chainName: getAllocationChainName(model, fallbackName),
    modelName: getAllocationChainName(model, fallbackName),
    status: normalizeValue(model.status),
    effectiveFrom: normalizeValue(model.effectiveFrom),
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

function loadAllocationSettingsModels() {
  try {
    const versioned = JSON.parse(
      localStorage.getItem(ALLOCATION_MODEL_VERSIONS_KEY) || "[]",
    );

    if (Array.isArray(versioned) && versioned.length > 0) {
      const normalized = versioned
        .map((model, index) =>
          normalizeAllocationSettingsModel(model, `Chain Allocation ${index + 1}`),
        )
        .filter(Boolean);

      const grouped = new Map();
      normalized.forEach((model) => {
        const chainName = getAllocationChainName(model, "Unassigned");
        const key = chainName.toLowerCase();
        const rows = grouped.get(key) || [];
        rows.push({ ...model, chainName, modelName: chainName });
        grouped.set(key, rows);
      });

      const activeModels = [];
      grouped.forEach((rows) => {
        const explicitActive = rows
          .filter((model) => normalizeValue(model.status) === "Active")
          .sort((left, right) =>
            normalizeValue(right.effectiveFrom).localeCompare(
              normalizeValue(left.effectiveFrom),
            ),
          )[0];

        const selected =
          explicitActive ||
          [...rows].sort((left, right) =>
            normalizeValue(right.effectiveFrom).localeCompare(
              normalizeValue(left.effectiveFrom),
            ),
          )[0];

        if (selected) {
          activeModels.push({ ...selected, status: "Active" });
        }
      });

      return activeModels.sort((left, right) =>
        getAllocationChainName(left).localeCompare(getAllocationChainName(right)),
      );
    }
  } catch {
    // Fall through to the legacy configuration.
  }

  try {
    const saved = JSON.parse(
      localStorage.getItem(ALLOCATION_SETTINGS_KEY) || "null",
    );
    const normalized = normalizeAllocationSettingsModel(
      saved,
      "Current Allocation",
    );
    return normalized ? [{ ...normalized, status: "Active" }] : [];
  } catch {
    return [];
  }
}

function findActiveAllocationSettings(models, chainName) {
  const normalizedChain = normalizeValue(chainName).toLowerCase();
  if (!normalizedChain) return null;
  return (
    models.find(
      (model) =>
        getAllocationChainName(model).toLowerCase() === normalizedChain &&
        normalizeValue(model.status) === "Active",
    ) || null
  );
}

function getAllocationPlan(count, planMode, settings, manualOtfAmount = "") {
  const numericCount = Math.max(1, Number(count) || 1);
  const slabs = Array.isArray(settings?.slabs) ? settings.slabs : [];
  const normalizedPlanMode = normalizeValue(planMode);

  const selectedSlab = normalizedPlanMode
    ? slabs.find(
        (slab) =>
          normalizeValue(slab?.label).toLowerCase() ===
          normalizedPlanMode.toLowerCase(),
      )
    : null;

  const countMatchedSlab = slabs.find((slab) => {
    const min = Number(slab?.min) || 1;
    const max =
      slab?.max === null || slab?.max === undefined || slab?.max === ""
        ? null
        : Number(slab.max);
    return numericCount >= min && (max === null || numericCount <= max);
  });

  const matchedSlab = selectedSlab || (!normalizedPlanMode ? countMatchedSlab : null);
  const hasConfiguredOtf = Boolean(normalizeValue(settings?.otfPricePerDevice));

  if (matchedSlab) {
    return {
      planMode: normalizeValue(matchedSlab.label),
      planFee: normalizeValue(matchedSlab.planFee),
      otfAmount: hasConfiguredOtf
        ? numericCount * Number(settings.otfPricePerDevice)
        : normalizeValue(manualOtfAmount),
    };
  }

  const legacyPlanMode =
    normalizedPlanMode === "More Than 3" ? "More Than 3" : "Up to 3";
  const legacyPlanFee =
    legacyPlanMode === "More Than 3"
      ? settings.moreThan3PlanFee
      : settings.upTo3PlanFee;

  return {
    planMode: normalizedPlanMode || legacyPlanMode,
    planFee: legacyPlanFee,
    otfAmount: hasConfiguredOtf
      ? numericCount * Number(settings.otfPricePerDevice)
      : normalizeValue(manualOtfAmount),
  };
}

function getAllocationPlanModeOptions(settings) {
  const slabs = Array.isArray(settings?.slabs) ? settings.slabs : [];
  const slabModes = slabs
    .map((slab) => normalizeValue(slab?.label))
    .filter(Boolean);

  if (slabModes.length > 0) return slabModes;

  const legacyModes = [];
  if (normalizeValue(settings?.upTo3PlanFee)) legacyModes.push("Up to 3");
  if (normalizeValue(settings?.moreThan3PlanFee)) legacyModes.push("More Than 3");
  return legacyModes;
}

function isValidRenewalDateValue(dateValue) {
  if (!dateValue) {
    return false;
  }

  if (!isValidDateInputValue(dateValue) || !isValidDateValue(dateValue)) {
    return false;
  }

  const year = Number(String(dateValue).slice(0, 4));
  return year >= 1000;
}

function formatDateValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

function parseDateInputValue(dateValue) {
  if (!isValidDateInputValue(dateValue) || !isValidDateValue(dateValue)) {
    return null;
  }

  const [year, month, day] = String(dateValue).split("-").map(Number);
  return new Date(Date.UTC(year, month - 1, day));
}

function isDateBefore(startDateValue, endDateValue) {
  const startDate = parseDateInputValue(startDateValue);
  const endDate = parseDateInputValue(endDateValue);

  if (!startDate || !endDate) {
    return false;
  }

  return endDate.getTime() < startDate.getTime();
}

function getAutoMouEndDate(startDateValue) {
  if (!isValidDateInputValue(startDateValue) || !isValidDateValue(startDateValue)) {
    return "";
  }

  const startDate = new Date(`${startDateValue}T00:00:00`);
  startDate.setFullYear(startDate.getFullYear() + 1);
  startDate.setDate(startDate.getDate() - 1);
  return formatDateValue(startDate);
}

function getRenewalPendingStartDate(endDateValue) {
  const endDate = parseDateInputValue(endDateValue);

  if (!endDate) {
    return null;
  }

  const renewalStart = new Date(endDate.getTime());
  const targetDay = renewalStart.getUTCDate();

  renewalStart.setUTCDate(1);
  renewalStart.setUTCMonth(renewalStart.getUTCMonth() - 1);

  const lastDayOfTargetMonth = new Date(
    Date.UTC(
      renewalStart.getUTCFullYear(),
      renewalStart.getUTCMonth() + 1,
      0,
    ),
  ).getUTCDate();

  renewalStart.setUTCDate(Math.min(targetDay, lastDayOfTargetMonth));
  return renewalStart;
}

function createDefaultPricingGroups() {
  return [
    {
      id: 1,
      groupName: "1",
      applicableSites: "",
      subscriptionMode: "Monthly",
      otfApplicable: "No",
      otfType: "",
      otfTaxMode: "",
      otfAmount: "0",
      subscriptionFee: "0",
    },
    {
      id: 2,
      groupName: "2",
      applicableSites: "",
      subscriptionMode: "Monthly",
      otfApplicable: "No",
      otfType: "",
      otfTaxMode: "",
      otfAmount: "0",
      subscriptionFee: "0",
    },
  ];
}

function MouSubscription({
  siteRows = [],
  onDataChange = () => {},
  onSave = () => {},
  onRequestAgreementChoice = () => {},
  agreementChoice = "",
  agreementChoiceConfirmed = false,
  requiredAgreementType = "MoU",
  resetSignal = 0,
  editGroupId = "",
  editGroupRecords = [],
  isEditingStage1 = false,
  onCancelEdit = () => {},
  addendumAvailable = false,
  addendumScope = "",
  addendumApplicableSiteIds = [],
}) {
  const [mouStatus, setMouStatus] = useState("Sent");
  const [mouSentDate, setMouSentDate] = useState("");
  const [mouReceivedDate, setMouReceivedDate] = useState("");
  const [mouStartDate, setMouStartDate] = useState("");
  const [mouEndDate, setMouEndDate] = useState("");
  const autoGeneratedMouEndDateRef = useRef("");
  const isMouEndDateManualRef = useRef(false);

  const [otfApplicable, setOtfApplicable] = useState("No");
  const [otfType, setOtfType] = useState("");
  const [otfTaxMode, setOtfTaxMode] = useState("");
  const [otfAmount, setOtfAmount] = useState("");

  const [subscriptionType, setSubscriptionType] = useState("Fixed");
  const [subscriptionMode, setSubscriptionMode] = useState("Monthly");
  const [subscriptionFee, setSubscriptionFee] = useState("");

  const [extensionRequired, setExtensionRequired] = useState("No");
  const [extensionRenewal, setExtensionRenewal] = useState("No");
  const [newMouStartDate, setNewMouStartDate] = useState("");
  const [newMouEndDate, setNewMouEndDate] = useState("");
  const [extensionRemarks, setExtensionRemarks] = useState("");

  const [incentiveApplicable, setIncentiveApplicable] = useState("Applicable");
  const [salesEmployeeName, setSalesEmployeeName] = useState("");
  const [companyId, setCompanyId] = useState("");
  const [pricingMethod, setPricingMethod] = useState("Common");
  const [pricingGroups, setPricingGroups] = useState(createDefaultPricingGroups);
  const [allocationSettingsModels, setAllocationSettingsModels] = useState(
    loadAllocationSettingsModels,
  );
  const [allocationChainNetwork, setAllocationChainNetwork] = useState("");
  const [allocationSettings, setAllocationSettings] = useState({
    ...DEFAULT_ALLOCATION_SETTINGS,
  });
  const [allocationRows, setAllocationRows] = useState([]);
  const [addendumEffectiveFrom, setAddendumEffectiveFrom] = useState("");
  const [addendumRemarks, setAddendumRemarks] = useState("");
  const [previousAgreementSnapshot, setPreviousAgreementSnapshot] = useState(null);

  const isAddendumSelected =
    normalizeValue(agreementChoice).toLowerCase() === "addendum";
  const activeEditSource = editGroupRecords[0] || {};
  const isAddendumWorkflow =
    (!isEditingStage1 && isAddendumSelected) ||
    normalizeValue(activeEditSource?.correctionType).toLowerCase() === "addendum";
  const isExistingCycleCommercialLocked =
    isEditingStage1 &&
    !isAddendumWorkflow &&
    (activeEditSource?.billingCommercialLocked === true ||
      ["Submitted to Billing Team", "Sent to Billing Team"].includes(
        normalizeValue(activeEditSource?.billingVerificationStatus),
      ) ||
      Boolean(normalizeValue(activeEditSource?.billingCommercialLockedAt)));
  const isReceived = mouStatus === "Received";
  const isSent = mouStatus === "Sent";
  const isOtfApplicable = otfApplicable === "Yes";
  const isVariableSubscription = subscriptionType === "Variable";
  const contextualMouDateLabel = isSent ? "MoU Sent Date" : "MoU Received Date";
  const contextualMouDateValue = isSent ? mouSentDate : mouReceivedDate;
  const contextualMouDateDisabled = !isSent && !isReceived;
  const billableSiteRows = siteRows.filter((row) => !isFoCSite(row));
  const completeSiteRows = siteRows.filter(
    (row) => normalizeValue(row.screenCode) && normalizeValue(row.screenName),
  );
  const isFoCOnlyStage1 =
    completeSiteRows.length > 0 &&
    completeSiteRows.every((row) => isFoCSite(row));
  const allocationEligibleSites = useMemo(
    () =>
      billableSiteRows
        .filter(
          (row) =>
            normalizeValue(row.screenCode) && normalizeValue(row.screenName),
        )
        .map((row) => ({
          screenCode: normalizeValue(row.screenCode),
          screenName: normalizeValue(row.screenName),
          complexCode: normalizeValue(row.complexCode || row.complexId),
          siteType: normalizeValue(row.siteType || row.groupType),
        })),
    [siteRows],
  );
  const allocationApplicable = useMemo(
    () =>
      allocationEligibleSites.some(
        (row) =>
          ["multiplex", "complex", "chain"].includes(row.siteType.toLowerCase()) ||
          Boolean(row.complexCode),
      ),
    [allocationEligibleSites],
  );
  const allocationChainNetworks = useMemo(
    () =>
      allocationSettingsModels
        .map((model) => getAllocationChainName(model))
        .filter(Boolean)
        .filter((chainName, index, values) =>
          values.findIndex(
            (value) => value.toLowerCase() === chainName.toLowerCase(),
          ) === index,
        )
        .sort((left, right) => left.localeCompare(right)),
    [allocationSettingsModels],
  );
  const isMouDateRangeInvalid =
    Boolean(mouStartDate && mouEndDate) &&
    isValidDateInputValue(mouStartDate) &&
    isValidDateInputValue(mouEndDate) &&
    isValidDateValue(mouStartDate) &&
    isValidDateValue(mouEndDate) &&
    isDateBefore(mouStartDate, mouEndDate);

  const pricingSiteLookups = useMemo(() => {
    const aliasToSite = new Map();
    const directToSite = new Map();
    const exactReferenceToSites = new Map();
    const screenNameToSites = new Map();
    let maxScreenNameWords = 1;
    const selectedRows = billableSiteRows.filter(
      (row) => row.screenCode?.trim() && row.screenName?.trim(),
    );

    function addLookup(map, key, siteEntry) {
      if (!key) {
        return;
      }

      const existing = map.get(key) || [];

      if (!existing.some((item) => item.screenCode === siteEntry.screenCode)) {
        map.set(key, [...existing, siteEntry]);
      }
    }

    selectedRows.forEach((row, index) => {
      const screenCode = normalizeValue(row.screenCode).toUpperCase();
      const screenName = normalizeValue(row.screenName).replace(/\s+/g, " ").toUpperCase();
      const erpScreenCode = normalizeValue(row.erpScreenCode || row.erpScreenName)
        .replace(/\s+/g, " ")
        .toUpperCase();
      const erpScreenName = normalizeValue(row.erpScreenName || row.screenName)
        .replace(/\s+/g, " ")
        .toUpperCase();
      const alias = `S${index + 1}`;
      const siteEntry = {
        screenCode,
        screenName,
        erpScreenCode,
        erpScreenName,
        alias,
      };

      aliasToSite.set(alias, siteEntry);

      if (screenCode) {
        directToSite.set(screenCode, siteEntry);
      }

      [screenCode, screenCode, erpScreenName, screenName]
        .filter(Boolean)
        .forEach((value) => {
          addLookup(exactReferenceToSites, value, siteEntry);
        });

      [screenName, erpScreenName]
        .filter(Boolean)
        .forEach((value) => {
          addLookup(screenNameToSites, value, siteEntry);
          maxScreenNameWords = Math.max(
            maxScreenNameWords,
            value.split(" ").length,
          );
        });
    });

    return {
      aliasToSite,
      directToSite,
      exactReferenceToSites,
      screenNameToSites,
      maxScreenNameWords,
    };
  }, [siteRows]);

  function normalizePricingReference(reference) {
    return normalizeValue(reference).replace(/\s+/g, " ").toUpperCase();
  }

  function getReferenceMatch(reference) {
    const normalizedReference = normalizePricingReference(reference);

    if (!normalizedReference) {
      return {
        error: "Empty reference.",
        siteIds: [],
      };
    }

    const rangeMatch = normalizedReference.match(/^S(\d+)\s*-\s*S(\d+)$/i);
    if (rangeMatch) {
      const start = Number(rangeMatch[1]);
      const end = Number(rangeMatch[2]);

      if (start > end) {
        return {
          error: `Invalid range ${reference.trim()}.`,
          siteIds: [],
        };
      }

      const siteIds = [];
      for (let index = start; index <= end; index += 1) {
        const rangeAlias = `S${index}`;
        const siteEntry = pricingSiteLookups.aliasToSite.get(rangeAlias);

        if (!siteEntry) {
          return {
            error: `Unknown reference: ${rangeAlias}.`,
            siteIds: [],
          };
        }

        if (!siteIds.includes(siteEntry.screenCode)) {
          siteIds.push(siteEntry.screenCode);
        }
      }

      return {
        siteIds,
      };
    }

    const aliasMatch = pricingSiteLookups.aliasToSite.get(normalizedReference);
    const directMatch = pricingSiteLookups.directToSite.get(normalizedReference);

    if (aliasMatch && directMatch && aliasMatch.screenCode !== directMatch.screenCode) {
      return {
        error: `${reference.trim()} is ambiguous. Use the Screen Code, S# alias, or screen name that uniquely identifies the site.`,
        siteIds: [],
      };
    }

    const directSite = aliasMatch || directMatch;
    if (directSite) {
      return {
        siteIds: [directSite.screenCode],
      };
    }

    const exactMatches =
      pricingSiteLookups.exactReferenceToSites.get(normalizedReference) || [];
    if (exactMatches.length === 1) {
      return {
        siteIds: [exactMatches[0].screenCode],
      };
    }

    if (exactMatches.length > 1) {
      return {
        error: `${normalizeValue(reference)} matches multiple selected sites. Use the Screen Code or S# alias.`,
        siteIds: [],
      };
    }

    return {
      error: `Unknown reference: ${normalizeValue(reference).replace(/\s+/g, " ")}.`,
      siteIds: [],
    };
  }

  function isUnknownReferenceError(error) {
    return String(error || "").startsWith("Unknown reference:");
  }

  function parsePricingGroupSites(applicableSites) {
    const rawValue = normalizeValue(applicableSites).replace(/\s+/g, " ");

    if (!rawValue) {
      return {
        siteIds: [],
        error: "",
      };
    }

    const seenSiteIds = new Set();
    const resolvedSiteIds = [];

    const segments = rawValue.split(",").map((segment) => normalizeValue(segment));

    for (const segment of segments) {
      if (!segment) {
        return {
          siteIds: [],
          error: "Applicable Sites / Screens cannot contain empty references.",
        };
      }

      const tokens = segment.split(" ").filter(Boolean);
      let index = 0;

      while (index < tokens.length) {
        const singleToken = tokens[index];
        const singleMatch = getReferenceMatch(singleToken);

        if (singleMatch.siteIds.length > 0) {
          for (const screenCode of singleMatch.siteIds) {
            if (seenSiteIds.has(screenCode)) {
              return {
                siteIds: [],
                error: `Duplicate reference found within the same group: ${singleToken}.`,
              };
            }

            seenSiteIds.add(screenCode);
            resolvedSiteIds.push(screenCode);
          }

          index += 1;
          continue;
        }

        if (singleMatch.error && !isUnknownReferenceError(singleMatch.error)) {
          return {
            siteIds: [],
            error: singleMatch.error,
          };
        }

        let phraseMatched = false;
        const maxWords = Math.min(
          pricingSiteLookups.maxScreenNameWords || 1,
          tokens.length - index,
        );

        for (let wordCount = maxWords; wordCount >= 2; wordCount -= 1) {
          const phrase = tokens.slice(index, index + wordCount).join(" ");
          const phraseMatch = getReferenceMatch(phrase);

          if (phraseMatch.siteIds.length === 0) {
            if (phraseMatch.error && !isUnknownReferenceError(phraseMatch.error)) {
              return {
                siteIds: [],
                error: phraseMatch.error,
              };
            }

            continue;
          }

          for (const screenCode of phraseMatch.siteIds) {
            if (seenSiteIds.has(screenCode)) {
              return {
                siteIds: [],
                error: `Duplicate reference found within the same group: ${phrase}.`,
              };
            }

            seenSiteIds.add(screenCode);
            resolvedSiteIds.push(screenCode);
          }

          index += wordCount;
          phraseMatched = true;
          break;
        }

        if (phraseMatched) {
          continue;
        }

        const unknownToken = singleToken;
        return {
          siteIds: [],
          error: `Unknown reference: ${unknownToken}.`,
        };
      }
    }

    return {
      siteIds: resolvedSiteIds,
      error: "",
    };
  }

  const pricingGroupsResolved = useMemo(
    () =>
      pricingGroups.map((group) => ({
        ...group,
        resolvedApplicableSiteIds: parsePricingGroupSites(group.applicableSites).siteIds,
      })),
    [pricingGroups, pricingSiteLookups],
  );

  const mouStartDateOnly = isValidRenewalDateValue(mouStartDate)
    ? new Date(`${mouStartDate}T00:00:00`)
    : null;
  const mouEndDateOnly = isValidRenewalDateValue(mouEndDate)
    ? new Date(`${mouEndDate}T00:00:00`)
    : null;
  const canEvaluateRenewalStatus =
    mouStartDateOnly !== null &&
    mouEndDateOnly !== null &&
    mouEndDateOnly >= mouStartDateOnly;
  const today = new Date();
  const todayOnly = new Date(
    today.getFullYear(),
    today.getMonth(),
    today.getDate(),
  );
  const isMouExpired = canEvaluateRenewalStatus
    ? mouEndDateOnly < todayOnly
    : false;
  const mouDaysRemaining = canEvaluateRenewalStatus
    ? Math.max(
        0,
        Math.ceil((mouEndDateOnly - todayOnly) / (1000 * 60 * 60 * 24)),
      )
    : null;
  const renewalPendingStartDate = canEvaluateRenewalStatus
    ? getRenewalPendingStartDate(mouEndDate)
    : null;
  const showRenewalWarning =
    canEvaluateRenewalStatus &&
    !isMouExpired &&
    renewalPendingStartDate !== null &&
    todayOnly >= renewalPendingStartDate;

  useEffect(() => {
    if (isExistingCycleCommercialLocked) {
      return;
    }

    const nextAutoEndDate = getAutoMouEndDate(mouStartDate);

    if (!nextAutoEndDate) {
      return;
    }

    const shouldAutoUpdate =
      mouEndDate === "" ||
      (!isMouEndDateManualRef.current &&
        mouEndDate === autoGeneratedMouEndDateRef.current);

    if (shouldAutoUpdate && mouEndDate !== nextAutoEndDate) {
      setMouEndDate(nextAutoEndDate);
    }

    if (!isMouEndDateManualRef.current) {
      autoGeneratedMouEndDateRef.current = nextAutoEndDate;
    }
  }, [mouEndDate, mouStartDate]);

  useEffect(() => {
    const refreshAllocationSettings = () => {
      const models = loadAllocationSettingsModels();
      setAllocationSettingsModels(models);

      if (allocationChainNetwork) {
        const selected = findActiveAllocationSettings(
          models,
          allocationChainNetwork,
        );
        if (selected) setAllocationSettings(selected);
      }
    };

    window.addEventListener(
      "billing-allocation-settings-updated",
      refreshAllocationSettings,
    );
    window.addEventListener("storage", refreshAllocationSettings);

    return () => {
      window.removeEventListener(
        "billing-allocation-settings-updated",
        refreshAllocationSettings,
      );
      window.removeEventListener("storage", refreshAllocationSettings);
    };
  }, [allocationChainNetwork]);

  useEffect(() => {
    if (pricingMethod !== PRICING_METHOD_ALLOCATION) return;

    setAllocationRows((currentRows) => {
      const byScreen = new Map(
        currentRows.map((row) => [row.screenCode, row]),
      );

      return allocationEligibleSites.map((site) => ({
        screenCode: site.screenCode,
        screenName: site.screenName,
        count: byScreen.get(site.screenCode)?.count || "1",
        planMode: byScreen.get(site.screenCode)?.planMode || "",
        otfAmount: byScreen.get(site.screenCode)?.otfAmount || "",
      }));
    });
  }, [pricingMethod, allocationEligibleSites]);

  useEffect(() => {
    if (resetSignal === 0) {
      return;
    }

    setMouStatus("Sent");
    setMouReceivedDate("");
    setMouStartDate("");
    setMouEndDate("");
    setOtfApplicable("No");
    setOtfType("");
    setOtfTaxMode("");
    setOtfAmount("");
    setSubscriptionType("Fixed");
    setSubscriptionMode("Monthly");
    setSubscriptionFee("");
    setExtensionRequired("No");
    setExtensionRenewal("No");
    setNewMouStartDate("");
    setNewMouEndDate("");
    setExtensionRemarks("");
    setIncentiveApplicable("Applicable");
    setSalesEmployeeName("");
    setCompanyId("");
    setPricingMethod("Common");
    setPricingGroups(createDefaultPricingGroups());
    setAllocationChainNetwork("");
    setAllocationSettings({ ...DEFAULT_ALLOCATION_SETTINGS });
    setAllocationRows([]);
    setAddendumEffectiveFrom("");
    setAddendumRemarks("");
    setPreviousAgreementSnapshot(null);
    autoGeneratedMouEndDateRef.current = "";
    isMouEndDateManualRef.current = false;
  }, [resetSignal]);

  useEffect(() => {
    if (!editGroupId) {
      return;
    }

    const sourceGroupRecord = editGroupRecords[0] || {};
    const pricingGroupsValue =
      Array.isArray(sourceGroupRecord.pricingGroups) &&
      sourceGroupRecord.pricingGroups.length > 0
        ? sourceGroupRecord.pricingGroups
        : createDefaultPricingGroups();
    const loadedStartDate = sourceGroupRecord.mouStartDate || "";
    const loadedEndDate = sourceGroupRecord.mouEndDate || "";
    const expectedEndDate = getAutoMouEndDate(loadedStartDate);

    setMouStatus(sourceGroupRecord.mouStatus || "Sent");
    setMouSentDate(sourceGroupRecord.mouSentDate || "");
    setMouReceivedDate(sourceGroupRecord.mouReceivedDate || "");
    setMouStartDate(loadedStartDate);
    setMouEndDate(loadedEndDate);
    setOtfApplicable(sourceGroupRecord.otfApplicable || "No");
    setOtfType(sourceGroupRecord.otfType || "");
    setOtfTaxMode(
      sourceGroupRecord.otfTaxMode ||
        (sourceGroupRecord.otfType === "Refundable"
          ? "Not Applicable"
          : ""),
    );
    setOtfAmount(sourceGroupRecord.otfAmount || "");
    setSubscriptionType(sourceGroupRecord.subscriptionType || "Fixed");
    setSubscriptionMode(sourceGroupRecord.subscriptionMode || "Monthly");
    setSubscriptionFee(sourceGroupRecord.subscriptionFee || "");
    setExtensionRequired(sourceGroupRecord.extensionRequired || "No");
    setExtensionRenewal(sourceGroupRecord.extensionRenewal || "No");
    setNewMouStartDate(sourceGroupRecord.newMouStartDate || "");
    setNewMouEndDate(sourceGroupRecord.newMouEndDate || "");
    setExtensionRemarks(sourceGroupRecord.extensionRemarks || "");
    setIncentiveApplicable(
      sourceGroupRecord.incentiveApplicable === false ||
      sourceGroupRecord.incentiveApplicable === "Not Applicable"
        ? "Not Applicable"
        : "Applicable",
    );
    setSalesEmployeeName(sourceGroupRecord.salesEmployeeName || "");
    setCompanyId(sourceGroupRecord.companyId || "");
    setPricingMethod(sourceGroupRecord.pricingMethod || "Common");
    setPricingGroups(pricingGroupsValue);

    const restoredAllocationSnapshot = normalizeAllocationSettingsModel(
      sourceGroupRecord.allocationSettingsSnapshot,
      "Saved Allocation",
    );
    const restoredChainNetwork =
      normalizeValue(sourceGroupRecord.allocationChainNetwork) ||
      normalizeValue(sourceGroupRecord.chainNetwork) ||
      getAllocationChainName(restoredAllocationSnapshot);
    setAllocationChainNetwork(restoredChainNetwork);
    setAllocationSettings(
      restoredAllocationSnapshot ||
        findActiveAllocationSettings(
          loadAllocationSettingsModels(),
          restoredChainNetwork,
        ) ||
        { ...DEFAULT_ALLOCATION_SETTINGS },
    );

    setAllocationRows(
      Array.isArray(sourceGroupRecord.allocationRows)
        ? sourceGroupRecord.allocationRows.map((row) => ({
            ...row,
            count: normalizeValue(row.count) || "1",
            planMode: normalizeValue(row.planMode),
          }))
        : [],
    );
    setAddendumEffectiveFrom(sourceGroupRecord.addendumEffectiveFrom || "");
    setAddendumRemarks(sourceGroupRecord.addendumRemarks || "");
    setPreviousAgreementSnapshot(
      sourceGroupRecord.previousAgreementSnapshot || {
        agreementType: sourceGroupRecord.agreementType || "MoU",
        mouStartDate: sourceGroupRecord.mouStartDate || "",
        mouEndDate: sourceGroupRecord.mouEndDate || "",
        subscriptionType: sourceGroupRecord.subscriptionType || "",
        subscriptionMode: sourceGroupRecord.subscriptionMode || "",
        subscriptionFee: sourceGroupRecord.subscriptionFee || "",
        otfApplicable: sourceGroupRecord.otfApplicable || "No",
        otfType: sourceGroupRecord.otfType || "",
        otfAmount: sourceGroupRecord.otfAmount || "",
        pricingMethod: sourceGroupRecord.pricingMethod || "Common",
        pricingGroups: Array.isArray(sourceGroupRecord.pricingGroups)
          ? sourceGroupRecord.pricingGroups.map((group) => ({ ...group }))
          : [],
        salesEmployeeName: sourceGroupRecord.salesEmployeeName || "",
        companyId: sourceGroupRecord.companyId || "",
      },
    );
    autoGeneratedMouEndDateRef.current = expectedEndDate;
    isMouEndDateManualRef.current =
      Boolean(loadedEndDate) && loadedEndDate !== expectedEndDate;
  }, [editGroupId, editGroupRecords]);

  useEffect(() => {
    if (subscriptionType === "Variable") {
      setSubscriptionMode("Monthly");
      setSubscriptionFee("");
    }
  }, [subscriptionType]);

  useEffect(() => {
    if (otfApplicable === "No") {
      setOtfType("");
      setOtfTaxMode("");
      setOtfAmount("");
    }
  }, [otfApplicable]);

  useEffect(() => {
    if (otfType === "Refundable") {
      setOtfTaxMode("Not Applicable");
      return;
    }

    if (otfTaxMode === "Not Applicable") {
      setOtfTaxMode("");
    }
  }, [otfType, otfTaxMode]);

  useEffect(() => {
    if (incentiveApplicable === "Not Applicable") {
      setSalesEmployeeName("");
      setCompanyId("");
    }
  }, [incentiveApplicable]);

  useEffect(() => {
    onDataChange({
      mouStatus: "",
      mouSentDate: "",
      mouReceivedDate: "",
      mouStartDate: "",
      mouEndDate: "",
      otfApplicable,
      otfType,
      otfTaxMode,
      otfAmount,
      subscriptionType,
      subscriptionMode,
      subscriptionFee,
      extensionRequired: "",
      extensionRenewal: "",
      newMouStartDate: "",
      newMouEndDate: "",
      extensionRemarks: "",
      incentiveApplicable,
      salesEmployeeName,
      companyId,
      pricingMethod,
      pricingGroups,
      pricingGroupsResolved,
      allocationChainNetwork,
      allocationModelId: normalizeValue(allocationSettings.modelId),
      allocationRows: allocationRows.map((row) => ({
        ...row,
        ...getAllocationPlan(row.count, row.planMode, allocationSettings, row.otfAmount),
      })),
      allocationSettingsSnapshot: { ...allocationSettings },
      agreementType: "",
      addendumEffectiveFrom: "",
      addendumRemarks: "",
      previousAgreementSnapshot: null,
      addendumScope: "",
      addendumApplicableSiteIds: [],
      requiresBillingDateReview: false,
    });
  }, [
    mouStatus,
    mouSentDate,
    mouReceivedDate,
    mouStartDate,
    mouEndDate,
    otfApplicable,
    otfType,
    otfTaxMode,
    otfAmount,
    subscriptionType,
    subscriptionMode,
    subscriptionFee,
    extensionRequired,
    extensionRenewal,
    newMouStartDate,
    newMouEndDate,
    extensionRemarks,
    incentiveApplicable,
    salesEmployeeName,
    companyId,
    pricingMethod,
    pricingGroups,
    pricingGroupsResolved,
    allocationChainNetwork,
    allocationRows,
    allocationSettings,
    agreementChoice,
    requiredAgreementType,
    isAddendumWorkflow,
    addendumEffectiveFrom,
    addendumRemarks,
    previousAgreementSnapshot,
    addendumScope,
    addendumApplicableSiteIds,
    onDataChange,
  ]);


  function handleSubscriptionTypeChange(event) {
    if (isExistingCycleCommercialLocked) {
      return;
    }

    const nextType = event.target.value;

    setSubscriptionType(nextType);

    if (nextType === "Variable") {
      setSubscriptionMode("Monthly");
      setSubscriptionFee("");
    }
  }

  function handleCompanyIdChange(event) {
    const numericValue = event.target.value.replace(/\D/g, "").slice(0, 5);
    setCompanyId(numericValue);
  }

  function handleValidatedDateChange(setter, { markEndDateManual = false } = {}) {
    return (event) => {
      const nextValue = event.target.value;

      if (!isValidDateInputValue(nextValue)) {
        alert("Year must contain exactly 4 digits.");
        return;
      }

      if (markEndDateManual) {
        isMouEndDateManualRef.current = true;
      }

      setter(nextValue);
    };
  }

  function handleMouStartDateChange(event) {
    const nextValue = event.target.value;

    if (!isValidDateInputValue(nextValue)) {
      alert("Year must contain exactly 4 digits.");
      return;
    }

    setMouStartDate(nextValue);

    if (!nextValue) {
      return;
    }

    const nextAutoEndDate = getAutoMouEndDate(nextValue);
    const shouldAutoUpdate =
      !isMouEndDateManualRef.current ||
      mouEndDate === "" ||
      mouEndDate === autoGeneratedMouEndDateRef.current;

    if (nextAutoEndDate && shouldAutoUpdate) {
      autoGeneratedMouEndDateRef.current = nextAutoEndDate;
      isMouEndDateManualRef.current = false;
      setMouEndDate(nextAutoEndDate);
    }
  }

  function handleContextualMouDateChange(event) {
    const nextValue = event.target.value;

    if (!isValidDateInputValue(nextValue)) {
      alert("Year must contain exactly 4 digits.");
      return;
    }

    if (isExistingCycleCommercialLocked) {
      if (isReceived) {
        setMouReceivedDate(nextValue);
      }
      return;
    }

    if (isSent) {
      setMouSentDate(nextValue);
      return;
    }

    if (isReceived) {
      setMouReceivedDate(nextValue);
    }
  }

  function handlePricingGroupChange(id, field, value) {
    if (isExistingCycleCommercialLocked) {
      return;
    }

    setPricingGroups((prev) =>
      prev.map((group) => {
        if (group.id !== id) return group;

        const nextGroup = { ...group, [field]: value };

        if (field === "otfApplicable" && value === "No") {
          nextGroup.otfType = "";
          nextGroup.otfTaxMode = "";
          nextGroup.otfAmount = "0";
        }

        if (field === "otfType") {
          nextGroup.otfTaxMode =
            value === "Refundable"
              ? "Not Applicable"
              : group.otfTaxMode === "Not Applicable"
                ? ""
                : group.otfTaxMode;
        }

        return nextGroup;
      }),
    );
  }

  function handleAddPricingGroup() {
    if (isExistingCycleCommercialLocked) {
      return;
    }

    setPricingGroups((prev) => [
      ...prev,
      {
        id: Date.now(),
        groupName: `${prev.length + 1}`,
        applicableSites: "",
        subscriptionMode: "Monthly",
        otfApplicable: "No",
        otfType: "",
        otfAmount: "0",
        subscriptionFee: "0",
      },
    ]);
  }

  function handleRemovePricingGroup(id) {
    if (isExistingCycleCommercialLocked) {
      return;
    }

    setPricingGroups((prev) => prev.filter((group) => group.id !== id));
  }

  function isBlankValue(value) {
    return value === "" || value === null || value === undefined || String(value).trim() === "";
  }

  function validatePricingGroups() {
    if (pricingMethod !== "Site-wise" || !allowSiteWisePricing) {
      return true;
    }

    const assignedSiteIds = new Set();
    const selectedSiteIds = completedSites.map((row) =>
      normalizeValue(row.screenCode).toUpperCase(),
    );

    for (const group of pricingGroupsResolved) {
      if (isBlankValue(group.applicableSites)) {
        alert(`Group ${group.groupName}: Enter Applicable Sites / Screens.`);
        return false;
      }

      if (isBlankValue(group.subscriptionMode)) {
        alert(`Group ${group.groupName}: Select Subscription Mode.`);
        return false;
      }

      if (isBlankValue(group.subscriptionFee)) {
        alert(`Group ${group.groupName}: Enter Subscription Fee.`);
        return false;
      }

      const { siteIds, error } = parsePricingGroupSites(group.applicableSites);

      if (error) {
        alert(error);
        return false;
      }

      if (group.otfApplicable === "Yes") {
        if (isBlankValue(group.otfType)) {
          alert(`Group ${group.groupName}: Select OTF Type.`);
          return false;
        }

        if (isBlankValue(group.otfAmount)) {
          alert(`Group ${group.groupName}: Enter OTF Amount.`);
          return false;
        }

        if (
          group.otfType === "Non-Refundable" &&
          isBlankValue(group.otfTaxMode)
        ) {
          alert(`Group ${group.groupName}: Select Tax.`);
          return false;
        }
      }

      for (const screenCode of siteIds) {
        if (assignedSiteIds.has(screenCode)) {
          alert(
            `${screenCode} is assigned to multiple pricing groups. Each selected site must be assigned once.`,
          );
          return false;
        }

        assignedSiteIds.add(screenCode);
      }
    }

    const unassignedSiteIds = selectedSiteIds.filter(
      (screenCode) => !assignedSiteIds.has(screenCode),
    );

    if (unassignedSiteIds.length > 0) {
      alert(
        `Each selected site must be assigned once. Missing: ${unassignedSiteIds.join(", ")}.`,
      );
      return false;
    }

    return true;
  }

  function validateCommonPricing() {
    if (isOtfApplicable) {
      if (isBlankValue(otfType)) {
        alert("Enter the OTF Type.");
        return false;
      }

      if (isBlankValue(otfAmount)) {
        alert("Enter the Common OTF Amount.");
        return false;
      }

      if (otfType === "Non-Refundable" && isBlankValue(otfTaxMode)) {
        alert("Select the OTF Tax.");
        return false;
      }
    }

    if (isBlankValue(subscriptionType)) {
      alert("Select the Subscription Type.");
      return false;
    }

    if (isBlankValue(subscriptionMode)) {
      alert("Select the Subscription Mode.");
      return false;
    }

    if (!isVariableSubscription && isBlankValue(subscriptionFee)) {
      alert("Enter the Common Subscription Fee Before GST.");
      return false;
    }

    return true;
  }

  function handleAllocationChainNetworkChange(event) {
    const nextChain = event.target.value;
    setAllocationChainNetwork(nextChain);

    const selectedSettings = findActiveAllocationSettings(
      allocationSettingsModels,
      nextChain,
    );
    setAllocationSettings(
      selectedSettings || { ...DEFAULT_ALLOCATION_SETTINGS },
    );
  }

  function handleAllocationCountChange(screenCode, value) {
    if (isExistingCycleCommercialLocked) return;

    const nextValue = String(value || "").replace(/\D/g, "");
    setAllocationRows((rows) =>
      rows.map((row) =>
        row.screenCode === screenCode
          ? { ...row, count: nextValue }
          : row,
      ),
    );
  }

  function handleAllocationPlanModeChange(screenCode, value) {
    if (isExistingCycleCommercialLocked) return;

    setAllocationRows((rows) =>
      rows.map((row) =>
        row.screenCode === screenCode
          ? { ...row, planMode: normalizeValue(value) }
          : row,
      ),
    );
  }

  function handleAllocationOtfAmountChange(screenCode, value) {
    if (isExistingCycleCommercialLocked) return;
    if (normalizeValue(allocationSettings?.otfPricePerDevice)) return;

    const nextValue = String(value || "").replace(/[^0-9.]/g, "");
    setAllocationRows((rows) =>
      rows.map((row) =>
        row.screenCode === screenCode
          ? { ...row, otfAmount: nextValue }
          : row,
      ),
    );
  }

  function handleRemoveAllocationRow(screenCode) {
    if (isExistingCycleCommercialLocked) return;
    setAllocationRows((rows) =>
      rows.filter((row) => row.screenCode !== screenCode),
    );
  }

  function handleAddAllocationSitePricing() {
    if (isExistingCycleCommercialLocked) return;

    const existing = new Set(
      allocationRows.map((row) => row.screenCode),
    );
    const nextSite = allocationEligibleSites.find(
      (site) => !existing.has(site.screenCode),
    );

    if (!nextSite) {
      alert("All Multiplex sites are already included in Allocation Model pricing.");
      return;
    }

    setAllocationRows((rows) => [
      ...rows,
      {
        screenCode: nextSite.screenCode,
        screenName: nextSite.screenName,
        count: "1",
        planMode: "",
        otfAmount: "",
      },
    ]);
  }

  function validateAllocationPricing() {
    if (pricingMethod !== PRICING_METHOD_ALLOCATION) return true;

    if (!allocationApplicable) {
      alert("Allocation Model is applicable only for Multiplex / Chain grouped sites.");
      return false;
    }

    if (!normalizeValue(allocationChainNetwork)) {
      alert("Select the Chain Network for Allocation Model pricing.");
      return false;
    }

    if (!normalizeValue(allocationSettings?.modelId)) {
      alert(
        `No Active Allocation is configured for ${allocationChainNetwork}. Configure and activate it in Settings first.`,
      );
      return false;
    }

    const configuredSlabs = Array.isArray(allocationSettings?.slabs)
      ? allocationSettings.slabs
      : [];
    const hasConfiguredPlanFees =
      configuredSlabs.length > 0
        ? configuredSlabs.every((slab) => normalizeValue(slab?.planFee))
        : Boolean(
            allocationSettings.upTo3PlanFee &&
              allocationSettings.moreThan3PlanFee,
          );

    if (!hasConfiguredPlanFees) {
      alert(
        "Complete the Allocation Plan Fee configuration in Settings before saving pricing.",
      );
      return false;
    }

    if (allocationRows.length === 0) {
      alert("Add at least one site to Allocation Model pricing.");
      return false;
    }

    const invalidRow = allocationRows.find(
      (row) => !Number(row.count) || Number(row.count) < 1,
    );

    if (invalidRow) {
      alert(`Enter a valid Device Count for ${invalidRow.screenCode}.`);
      return false;
    }

    const rowWithoutPlanMode = allocationRows.find(
      (row) => !normalizeValue(row.planMode),
    );

    if (rowWithoutPlanMode) {
      alert(`Select a Plan Mode for ${rowWithoutPlanMode.screenCode}.`);
      return false;
    }

    const hasConfiguredOtf = Boolean(
      normalizeValue(allocationSettings?.otfPricePerDevice),
    );

    if (!hasConfiguredOtf) {
      const invalidOtfRow = allocationRows.find((row) => {
        const value = normalizeValue(row.otfAmount);
        return value === "" || !Number.isFinite(Number(value)) || Number(value) < 0;
      });

      if (invalidOtfRow) {
        alert(
          `Enter the OTF Amount for ${invalidOtfRow.screenCode}, because OTF is not configured in Settings.`,
        );
        return false;
      }
    }

    return true;
  }

  function handleSavePricing() {
    if (isExistingCycleCommercialLocked) {
      alert(
        "Commercial pricing is locked because this cycle was already shared with the Billing Team.",
      );
      return false;
    }

    if (pricingMethod === "Common") {
      if (!validateCommonPricing()) {
        return false;
      }
    } else if (pricingMethod === PRICING_METHOD_ALLOCATION) {
      if (!validateAllocationPricing()) {
        return false;
      }
    } else if (!validatePricingGroups()) {
      return false;
    }

    alert("Pricing details saved successfully.");
    return true;
  }

  function handleSave() {
    if (isFoCOnlyStage1) {
      return true;
    }

    if (pricingMethod === "Common") {
      if (!validateCommonPricing()) {
        return false;
      }
    } else if (pricingMethod === PRICING_METHOD_ALLOCATION) {
      if (!validateAllocationPricing()) {
        return false;
      }
    } else if (!validatePricingGroups()) {
      return false;
    }

    if (incentiveApplicable === "Applicable") {
      if (!salesEmployeeName.trim()) {
        alert("Enter the Sales Employee Name.");
        return false;
      }

      if (companyId.length !== 5) {
        alert("Company ID must contain exactly 5 digits.");
        return false;
      }
    }

    return true;
  }

  function handleSaveClick() {
    if (handleSave()) {
      onSave();
    }
  }

const completedSites = billableSiteRows.filter(
  (row) => row.screenCode?.trim() && row.screenName?.trim(),
);

const siteCount = completedSites.length;
const allowSiteWisePricing = siteCount > 1;
const sitePricingRows = completedSites.map((row, index) => {
  const key = row.screenCode?.trim() || row.screenName?.trim() || `site-${index}`;
  const label = row.screenCode && row.screenName
    ? `${row.screenCode} / ${row.screenName}`
    : row.screenCode || row.screenName || `Site ${index + 1}`;

  return { key, label };
});

return (
  <section className="mou-subscription">
      <style>{`
        .mou-subscription__text-action {
          background: transparent;
          border: 0;
          color: inherit;
          opacity: 0.78;
          padding: 8px 10px;
          cursor: pointer;
          text-decoration: underline;
          text-underline-offset: 3px;
        }

        .mou-subscription__text-action:hover {
          opacity: 1;
        }

        .mou-subscription__status-badge {
          display: inline-flex;
          align-items: center;
          min-height: 36px;
          padding: 7px 12px;
          border-radius: 999px;
          border: 1px solid rgba(148, 163, 184, 0.35);
          background: rgba(148, 163, 184, 0.10);
          font-weight: 700;
          white-space: nowrap;
        }

        .mou-subscription__button--primary {
          font-weight: 800;
        }
      `}</style>
      <div className="mou-subscription__section">
        <h4>Subscription</h4>

        <div
          className="mou-subscription__pricing-options"
          style={{ gridTemplateColumns: "repeat(3, minmax(0, 1fr))" }}
        >
          <label className="mou-subscription__pricing-option">
            <input
              type="radio"
              name="pricingMethod"
              value="Common"
              checked={pricingMethod === "Common"}
              onChange={(event) => setPricingMethod(event.target.value)}
              disabled={isExistingCycleCommercialLocked}
            />

            <span>
              <strong>Common for selected sites</strong>
              <small>
                Common OTF and subscription amount for selected sites
              </small>
            </span>
          </label>

          <label
            className={`mou-subscription__pricing-option ${
              !allowSiteWisePricing
                ? "mou-subscription__pricing-option--disabled"
                : ""
            }`}
          >
            <input
              type="radio"
              name="pricingMethod"
              value="Site-wise"
              checked={pricingMethod === "Site-wise"}
              onChange={(event) => setPricingMethod(event.target.value)}
              disabled={
                !allowSiteWisePricing || isExistingCycleCommercialLocked
              }
            />

            <span>
              <strong>Site / Screen-wise</strong>
              <small>
                Enter separate or grouped pricing for the selected Screen Codes.
              </small>
            </span>
          </label>

          <label
            className={`mou-subscription__pricing-option ${
              !allocationApplicable
                ? "mou-subscription__pricing-option--disabled"
                : ""
            }`}
          >
            <input
              type="radio"
              name="pricingMethod"
              value={PRICING_METHOD_ALLOCATION}
              checked={pricingMethod === PRICING_METHOD_ALLOCATION}
              onChange={(event) => setPricingMethod(event.target.value)}
              disabled={
                !allocationApplicable || isExistingCycleCommercialLocked
              }
            />

            <span>
              <strong>Allocation Model</strong>
              <small>Pre-Set Screen Pricing</small>
            </span>
          </label>
        </div>

        {pricingMethod === "Common" && (
          <div className="mou-subscription__site-pricing-panel">
            <div className="mou-subscription__common-otf-block">
              <h4 style={{ margin: "0 0 14px" }}>OTF</h4>
              <div
                className="mou-subscription__common-otf-grid"
                style={{ gridTemplateColumns: "repeat(2, minmax(0, 1fr))" }}
              >
                <label className="mou-subscription__field">
                  <span>OTF Applicable</span>
                  <select
                    value={otfApplicable}
                    onChange={(event) => setOtfApplicable(event.target.value)}
                    disabled={isExistingCycleCommercialLocked}
                  >
                    <option value="No">No</option>
                    <option value="Yes">Yes</option>
                  </select>
                </label>

                <label className="mou-subscription__field">
                  <span>OTF Type</span>
                  <select
                    value={otfType}
                    onChange={(event) => setOtfType(event.target.value)}
                    disabled={
                      otfApplicable === "No" || isExistingCycleCommercialLocked
                    }
                  >
                    <option value="">Select type</option>
                    <option value="Refundable">Refundable</option>
                    <option value="Non-Refundable">Non-Refundable</option>
                  </select>
                </label>

                <label className="mou-subscription__field mou-subscription__common-otf-amount">
                  <span>OTF Amount</span>
                  <input
                    type="text"
                    inputMode="decimal"
                    value={otfAmount}
                    onChange={(event) => setOtfAmount(sanitizeAmount(event.target.value))}
                    disabled={
                      otfApplicable === "No" ||
                      !otfType ||
                      isExistingCycleCommercialLocked
                    }
                    placeholder="Enter common OTF amount"
                  />
                </label>

                <label className="mou-subscription__field">
                  <span>Tax</span>
                  <select
                    value={otfTaxMode}
                    onChange={(event) => setOtfTaxMode(event.target.value)}
                    disabled={
                      otfApplicable === "No" ||
                      !otfType ||
                      otfType === "Refundable" ||
                      isExistingCycleCommercialLocked
                    }
                  >
                    <option value="">Select tax</option>
                    {otfType === "Refundable" && (
                      <option value="Not Applicable">Not Applicable</option>
                    )}
                    <option value="Including Tax">Including Tax</option>
                    <option value="Excluding Tax">Excluding Tax</option>
                  </select>
                </label>
              </div>
            </div>
          </div>
        )}

        {pricingMethod === PRICING_METHOD_ALLOCATION &&
          allocationApplicable && (
            <div
              className="mou-subscription__workspace"
              style={{ width: "100%", maxWidth: "100%", minWidth: 0 }}
            >
              <div className="mou-subscription__workspace-header">
                <div>
                  <h4>Allocation Model</h4>
                  <small>
                    Select the Chain Network. Plan Fee is fetched from Settings. OTF Amount is fetched when configured in Settings; otherwise enter it here.
                  </small>
                </div>
              </div>

              <div className="mou-subscription__field-grid" style={{ marginBottom: 14 }}>
                <label className="mou-subscription__field">
                  <span>Chain Network *</span>
                  <select
                    value={allocationChainNetwork}
                    onChange={handleAllocationChainNetworkChange}
                    disabled={isExistingCycleCommercialLocked}
                  >
                    <option value="">Select Chain Network</option>
                    {allocationChainNetworks.map((chainName) => (
                      <option key={chainName} value={chainName}>
                        {chainName}
                      </option>
                    ))}
                  </select>
                </label>
              </div>

              {allocationChainNetworks.length === 0 ? (
                <p className="mou-subscription__helper">
                  No Active Chain Allocation is available. Create and Activate a Chain Allocation in Settings first.
                </p>
              ) : null}

              <div
                className="mou-subscription__pricing-table-wrapper"
                style={{ overflowX: "auto" }}
              >
                <table
                  className="mou-subscription__pricing-table"
                  style={{ width: "100%", minWidth: "760px", tableLayout: "fixed" }}
                >
                  <colgroup>
                    <col style={{ width: "34%" }} />
                    <col style={{ width: "13%" }} />
                    <col style={{ width: "16%" }} />
                    <col style={{ width: "14%" }} />
                    <col style={{ width: "15%" }} />
                    <col style={{ width: "8%" }} />
                  </colgroup>
                  <thead>
                    <tr>
                      <th>Screen</th>
                      <th>Device Count</th>
                      <th>OTF Amount</th>
                      <th>Plan Mode</th>
                      <th>Plan Fee</th>
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {allocationRows.map((row) => {
                      const calculated = getAllocationPlan(
                        row.count,
                        row.planMode,
                        allocationSettings,
                        row.otfAmount,
                      );

                      return (
                        <tr key={row.screenCode}>
                          <td
                            title={
                              row.screenCode || row.screenName
                                ? `${row.screenCode || ""}${
                                    row.screenCode && row.screenName ? " / " : ""
                                  }${row.screenName || ""}`
                                : undefined
                            }
                          >
                            <span className="table-cell-ellipsis">{row.screenCode} / {row.screenName}</span>
                          </td>
                          <td>
                            <input
                              type="number"
                              min="1"
                              step="1"
                              value={row.count}
                              onChange={(event) =>
                                handleAllocationCountChange(
                                  row.screenCode,
                                  event.target.value,
                                )
                              }
                              disabled={isExistingCycleCommercialLocked}
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              inputMode="decimal"
                              value={
                                normalizeValue(allocationSettings?.otfPricePerDevice)
                                  ? Number(calculated.otfAmount || 0).toLocaleString("en-IN")
                                  : row.otfAmount || ""
                              }
                              onChange={(event) =>
                                handleAllocationOtfAmountChange(
                                  row.screenCode,
                                  event.target.value,
                                )
                              }
                              disabled={
                                Boolean(normalizeValue(allocationSettings?.otfPricePerDevice)) ||
                                isExistingCycleCommercialLocked
                              }
                              placeholder={
                                normalizeValue(allocationSettings?.otfPricePerDevice)
                                  ? "Fetched from Settings"
                                  : "Enter OTF Amount"
                              }
                              title={
                                normalizeValue(allocationSettings?.otfPricePerDevice)
                                  ? "OTF Amount is calculated from the OTF configured in Settings."
                                  : "OTF is not configured in Settings, so enter the OTF Amount here."
                              }
                            />
                          </td>
                          <td>
                            <select
                              value={row.planMode || ""}
                              onChange={(event) =>
                                handleAllocationPlanModeChange(
                                  row.screenCode,
                                  event.target.value,
                                )
                              }
                              disabled={isExistingCycleCommercialLocked}
                            >
                              <option value="">Select Plan Mode</option>
                              {getAllocationPlanModeOptions(allocationSettings).map(
                                (mode) => (
                                  <option key={mode} value={mode}>
                                    {mode}
                                  </option>
                                ),
                              )}
                            </select>
                          </td>
                          <td>
                            ₹{Number(calculated.planFee || 0).toLocaleString("en-IN")}
                          </td>
                          <td>
                            <button
                              type="button"
                              className="mou-subscription__pricing-remove"
                              onClick={() =>
                                handleRemoveAllocationRow(row.screenCode)
                              }
                              disabled={isExistingCycleCommercialLocked}
                            >
                              ×
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              <div className="mou-subscription__workspace-actions">
                <button
                  type="button"
                  className="mou-subscription__button mou-subscription__button--secondary"
                  onClick={handleAddAllocationSitePricing}
                  disabled={isExistingCycleCommercialLocked}
                >
                  + Add Site Pricing
                </button>
              </div>
            </div>
          )}

        {pricingMethod === "Site-wise" && allowSiteWisePricing && (
          <div
            className="mou-subscription__workspace"
            style={{ width: "100%", maxWidth: "100%", minWidth: 0 }}
          >
            <div className="mou-subscription__workspace-header">
              <h4>Site / Screen-wise Pricing</h4>
            </div>

            <div
              className="mou-subscription__pricing-table-wrapper"
              style={{
                width: "100%",
                maxWidth: "100%",
                minWidth: 0,
                overflowX: "auto",
              }}
            >
              <table
                className="mou-subscription__pricing-table"
                style={{
                  width: "100%",
                  minWidth: "900px",
                  tableLayout: "fixed",
                }}
              >
                <colgroup>
                  <col
                    className="mou-subscription__pricing-col mou-subscription__pricing-col--group"
                    style={{ width: "7%" }}
                  />
                  <col
                    className="mou-subscription__pricing-col mou-subscription__pricing-col--sites"
                    style={{ width: "19%" }}
                  />
                  <col
                    className="mou-subscription__pricing-col mou-subscription__pricing-col--mode"
                    style={{ width: "12%" }}
                  />
                  <col
                    className="mou-subscription__pricing-col mou-subscription__pricing-col--otf-applicable"
                    style={{ width: "9%" }}
                  />
                  <col
                    className="mou-subscription__pricing-col mou-subscription__pricing-col--otf-type"
                    style={{ width: "13%" }}
                  />
                  <col
                    className="mou-subscription__pricing-col mou-subscription__pricing-col--otf-amount"
                    style={{ width: "13%" }}
                  />
                  <col
                    className="mou-subscription__pricing-col mou-subscription__pricing-col--subscription-fee"
                    style={{ width: "17%" }}
                  />
                  <col
                    className="mou-subscription__pricing-col mou-subscription__pricing-col--action"
                    style={{ width: "10%" }}
                  />
                </colgroup>
                <thead>
                  <tr>
                    <th>Group</th>
                    <th>Screens</th>
                    <th>Mode</th>
                    <th>OTF</th>
                    <th>OTF Type</th>
                    <th>Tax</th>
                    <th>OTF Amount</th>
                    <th>Plan Fee</th>
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {pricingGroups.map((group) => (
                    <tr key={group.id}>
                      <td>{group.groupName}</td>
                      <td>
                        <input
                          type="text"
                          value={group.applicableSites}
                          onChange={(event) =>
                            handlePricingGroupChange(
                              group.id,
                              "applicableSites",
                              event.target.value,
                            )
                          }
                          placeholder="Example: S1, S2 or Geeth, Susat"
                          disabled={isExistingCycleCommercialLocked}
                        />
                      </td>
                      <td>
                        <select
                          value={group.subscriptionMode}
                          onChange={(event) =>
                            handlePricingGroupChange(
                              group.id,
                              "subscriptionMode",
                              event.target.value,
                            )
                          }
                        >
                          <option value="Monthly">Monthly</option>
                          <option value="Quarterly">Quarterly</option>
                          <option value="Half Yearly">Half Yearly</option>
                          <option value="Yearly">Yearly</option>
                        </select>
                      </td>
                      <td>
                        <select
                          value={group.otfApplicable}
                          onChange={(event) =>
                            handlePricingGroupChange(
                              group.id,
                              "otfApplicable",
                              event.target.value,
                            )
                          }
                        >
                          <option value="No">No</option>
                          <option value="Yes">Yes</option>
                        </select>
                      </td>
                      <td>
                        <select
                          value={group.otfType}
                          onChange={(event) =>
                            handlePricingGroupChange(
                              group.id,
                              "otfType",
                              event.target.value,
                            )
                          }
                          disabled={
                            group.otfApplicable === "No" ||
                            isExistingCycleCommercialLocked
                          }
                        >
                          <option value="">Select type</option>
                          <option value="Refundable">Refundable</option>
                          <option value="Non-Refundable">
                            Non-Refundable
                          </option>
                        </select>
                      </td>
                      <td>
                        <select
                          value={group.otfTaxMode || ""}
                          onChange={(event) =>
                            handlePricingGroupChange(
                              group.id,
                              "otfTaxMode",
                              event.target.value,
                            )
                          }
                          disabled={
                            group.otfApplicable === "No" ||
                            !group.otfType ||
                            group.otfType === "Refundable" ||
                            isExistingCycleCommercialLocked
                          }
                        >
                          <option value="">Select tax</option>
                          {group.otfType === "Refundable" && (
                            <option value="Not Applicable">
                              Not Applicable
                            </option>
                          )}
                          <option value="Including Tax">Including Tax</option>
                          <option value="Excluding Tax">Excluding Tax</option>
                        </select>
                      </td>
                      <td>
                        <input
                          type="text"
                          inputMode="decimal"
                          className="mou-subscription__pricing-input mou-subscription__pricing-input--narrow"
                          value={group.otfAmount}
                          onChange={(event) =>
                            handlePricingGroupChange(
                              group.id,
                              "otfAmount",
                              sanitizeAmount(event.target.value),
                            )
                          }
                          disabled={
                            group.otfApplicable === "No" ||
                            !normalizeValue(group.otfType) ||
                            isExistingCycleCommercialLocked
                          }
                        />
                      </td>
                      <td>
                        <input
                          type="text"
                          inputMode="decimal"
                          className="mou-subscription__pricing-input mou-subscription__pricing-input--narrow"
                          value={group.subscriptionFee}
                          onChange={(event) =>
                            handlePricingGroupChange(
                              group.id,
                              "subscriptionFee",
                              sanitizeAmount(event.target.value),
                            )
                          }
                          disabled={isExistingCycleCommercialLocked}
                        />
                      </td>
                      <td>
                        <button
                          type="button"
                          className="mou-subscription__pricing-remove"
                          onClick={() => handleRemovePricingGroup(group.id)}
                          disabled={isExistingCycleCommercialLocked}
                          aria-label="Remove pricing group"
                          title="Remove pricing group"
                        >
                          ×
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="mou-subscription__workspace-actions">
              <button
                type="button"
                className="mou-subscription__button mou-subscription__button--secondary"
                onClick={handleAddPricingGroup}
                disabled={isExistingCycleCommercialLocked}
              >
                + Add Pricing Group
              </button>

              <div className="mou-subscription__workspace-actions-right">
                <button
                  type="button"
                  className="mou-subscription__text-action"
                >
                  Refresh Selected Sites
                </button>
                <button
                  type="button"
                  className="mou-subscription__text-action"
                >
                  Return to Edit
                </button>
                <button
                  type="button"
                  className="mou-subscription__button"
                  onClick={handleSavePricing}
                  disabled={isExistingCycleCommercialLocked}
                >
                  {isExistingCycleCommercialLocked
                    ? "Pricing Locked"
                    : "Save Pricing"}
                </button>
              </div>
            </div>
          </div>
        )}

        <div className="mou-subscription__message">
          {siteCount === 1 &&
            "Standalone or single-site entry uses Common pricing."}

          {siteCount >= 2 &&
            siteCount <= 4 &&
            "Site-wise pricing can be entered separately for each Screen Code."}

          {siteCount >= 5 &&
            "Five or more sites can use Common, Group A, Group B or Group C pricing groups."}
        </div>
      </div>

      {pricingMethod === "Common" && (
        <div className="mou-subscription__section">
          <h4>Subscription Details</h4>

          <div
            className="mou-subscription__form-grid"
            style={{ gridTemplateColumns: "repeat(3, minmax(0, 1fr))" }}
          >
            <label className="mou-subscription__field">
              <span>Subscription Type</span>
              <select
                value={subscriptionType}
                onChange={handleSubscriptionTypeChange}
                disabled={isExistingCycleCommercialLocked}
              >
                <option value="Fixed">Fixed</option>
                <option value="Variable">Variable</option>
              </select>
            </label>

            <label className="mou-subscription__field">
              <span>Subscription Mode</span>
              <select
                value={subscriptionMode}
                onChange={(event) => setSubscriptionMode(event.target.value)}
                disabled={
                  isVariableSubscription || isExistingCycleCommercialLocked
                }
              >
                <option value="Monthly">Monthly</option>
                <option value="Half-Yearly">Half-Yearly</option>
                <option value="Annual">Annual</option>
              </select>
            </label>

            <label className="mou-subscription__field">
              <span>Subscription Fee Before GST</span>
              <input
                type="text"
                inputMode="decimal"
                value={subscriptionFee}
                onChange={(event) => setSubscriptionFee(sanitizeAmount(event.target.value))}
                placeholder={
                  isVariableSubscription
                    ? "Captured month-wise in Stage 3"
                    : "Enter subscription fee"
                }
                disabled={
                  isVariableSubscription || isExistingCycleCommercialLocked
                }
              />
            </label>
          </div>

          <div
            className="mou-subscription__message"
            style={{ marginTop: 10, whiteSpace: "nowrap" }}
          >
            {isVariableSubscription
              ? "Monthly mode only."
              : "Fixed subscription supports Monthly, Half-Yearly and Annual modes."}
          </div>
        </div>
      )}

      <div className="mou-subscription__section">
        <h4>Incentive Beneficiary</h4>

        <div
          className="mou-subscription__form-grid"
          style={{ gridTemplateColumns: "repeat(3, minmax(0, 1fr))" }}
        >
          <label className="mou-subscription__field">
            <span>Employee Name</span>
            <input
              type="text"
              value={salesEmployeeName}
              onChange={(event) => setSalesEmployeeName(event.target.value)}
              placeholder={
                incentiveApplicable === "Not Applicable"
                  ? "Not Applicable"
                  : "Enter salesperson name"
              }
              disabled={incentiveApplicable === "Not Applicable"}
            />
          </label>

          <label className="mou-subscription__field">
            <span>
              Employee Number <small>(5 digits)</small>
            </span>
            <input
              type="text"
              inputMode="numeric"
              value={companyId}
              onChange={handleCompanyIdChange}
              placeholder={
                incentiveApplicable === "Not Applicable"
                  ? "Not Applicable"
                  : "Exactly 5 digits"
              }
              disabled={incentiveApplicable === "Not Applicable"}
            />
          </label>

          <label className="mou-subscription__field">
            <span>Incentive</span>
            <select
              value={incentiveApplicable}
              onChange={(event) => setIncentiveApplicable(event.target.value)}
            >
              <option value="Applicable">Applicable</option>
              <option value="Not Applicable">Not Applicable</option>
            </select>
          </label>
        </div>
      </div>

      <div className="mou-subscription__actions">
        <button
          type="button"
          className="mou-subscription__button mou-subscription__button--primary"
          onClick={handleSaveClick}
        >
          {isAddendumWorkflow
            ? "Save Addendum"
            : isEditingStage1
              ? "Update 1B Allocation"
              : "Save Stage 1"}
        </button>

        {isEditingStage1 && (
          <button
            type="button"
            className="mou-subscription__button mou-subscription__button--secondary"
            onClick={onCancelEdit}
          >
            Cancel Edit
          </button>
        )}
      </div>
    </section>
  );
}

export default MouSubscription;
