export const BILLING_API_BASE =
  import.meta.env.VITE_BILLING_API_BASE || "http://localhost:4100/api";

function normalizeValue(value) {
  return String(value ?? "").trim();
}

export async function billingApiRequest(path, options = {}) {
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
        ? `${baseMessage} ${validationMessages.join(" ")}`
        : baseMessage,
    );
  }

  if (
    typeof window !== "undefined" &&
    String(options.method || "GET").toUpperCase() !== "GET"
  ) {
    window.dispatchEvent(new CustomEvent("billing-summary-updated"));
  }

  return body?.data ?? body;
}

export async function getStage3BillingSummary() {
  return billingApiRequest("/stage3/allocation-reallocation/billing-summary");
}

const LEGACY_CONFIG_KEY = "billing_inc_allocation_model_config";
const LEGACY_VERSIONS_KEY = "billing_inc_allocation_model_versions";
const LEGACY_DRAFTS_KEY = "billing_inc_allocation_model_drafts";
const MIGRATION_MARKER_KEY = "billing_inc_allocation_backend_migrated_v1";

function parseJsonStorage(key, fallback) {
  if (typeof window === "undefined") return fallback;
  try {
    return JSON.parse(window.localStorage.getItem(key) || JSON.stringify(fallback));
  } catch {
    return fallback;
  }
}

function getLegacyIdentity(model, fallback = "") {
  const purpose = normalizeValue(model?.allocationPurpose || model?.purpose);
  const raw = normalizeValue(
    model?.chainName ||
      model?.networkName ||
      model?.name ||
      model?.modelName ||
      fallback,
  );
  let chainName = raw;
  let allocationPurpose = purpose;
  const stage3Match = raw.match(/^(.*?)\s*(?:[-–—·:]\s*)?Stage\s*3$/i);
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

  return {
    chainName,
    allocationPurpose,
    displayName: [chainName, allocationPurpose].filter(Boolean).join(" ") || raw,
  };
}

function normalizeLegacyPackages(model) {
  const packages = Array.isArray(model?.packages)
    ? model.packages
    : Array.isArray(model?.slabs)
      ? model.slabs
      : [];

  if (packages.length > 0) {
    return packages.map((item, index) => ({
      packageId: normalizeValue(item?.packageId || item?.slabId) || `legacy-${index + 1}`,
      packageName: normalizeValue(item?.packageName || item?.label),
      min:
        item?.min ?? item?.minQty ?? "",
      max:
        item?.max ?? item?.maxQty ?? "",
      noLimit:
        Boolean(item?.noLimit) ||
        item?.max === null ||
        item?.maxQty === null,
      planFee: normalizeValue(item?.planFee),
    }));
  }

  const legacy = [];
  if (normalizeValue(model?.upTo3PlanFee)) {
    legacy.push({
      packageId: "legacy-up-to-3",
      packageName: "Up to 3",
      min: 1,
      max: 3,
      noLimit: false,
      planFee: normalizeValue(model.upTo3PlanFee),
    });
  }
  if (normalizeValue(model?.moreThan3PlanFee)) {
    legacy.push({
      packageId: "legacy-more-than-3",
      packageName: "More Than 3",
      min: 4,
      max: "",
      noLimit: true,
      planFee: normalizeValue(model.moreThan3PlanFee),
    });
  }
  return legacy;
}

function normalizeLegacyModel(model, fallback = "Legacy Allocation") {
  if (!model || typeof model !== "object") return null;
  const identity = getLegacyIdentity(model, fallback);
  return {
    legacyId: normalizeValue(model.modelId || model.id || model.draftId),
    chainName: identity.chainName,
    allocationPurpose: identity.allocationPurpose,
    name: identity.displayName,
    effectiveFrom: normalizeValue(model.effectiveFrom),
    effectiveTo: normalizeValue(model.effectiveTo),
    otfPricePerDevice: normalizeValue(model.otfPricePerDevice),
    activationReason: normalizeValue(model.activationReason),
    remarks: normalizeValue(model.remarks),
    sourceModelId: normalizeValue(model.sourceModelId || model.sourceAllocationId),
    status: normalizeValue(model.status) || "Active",
    packages: normalizeLegacyPackages(model),
    savedAt: normalizeValue(model.savedAt || model.updatedAt || model.createdAt),
  };
}

function isBackendActivatable(model) {
  if (!model?.chainName) return false;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(normalizeValue(model.effectiveFrom))) return false;
  if (!Array.isArray(model.packages) || model.packages.length === 0) return false;

  return model.packages.every((pkg) => {
    const min = Number(pkg.min);
    const max = pkg.noLimit ? null : Number(pkg.max);
    const fee = Number(pkg.planFee);
    return (
      normalizeValue(pkg.packageName) &&
      Number.isInteger(min) &&
      min >= 1 &&
      (pkg.noLimit || (Number.isInteger(max) && max >= min)) &&
      Number.isFinite(fee) &&
      fee >= 0
    );
  });
}

function modelToBackendPayload(model) {
  return {
    chainName: model.chainName || null,
    allocationPurpose: model.allocationPurpose || null,
    name: model.name || null,
    displayName: model.name || null,
    effectiveFrom: model.effectiveFrom || null,
    otfPricePerDevice:
      model.otfPricePerDevice === "" ? null : model.otfPricePerDevice,
    sourceModelId: model.sourceModelId || null,
    activationReason: model.activationReason || null,
    remarks: model.remarks || null,
    packages: (model.packages || []).map((pkg) => ({
      packageId: pkg.packageId || null,
      packageName: pkg.packageName || null,
      min: pkg.min === "" ? null : pkg.min,
      max: pkg.noLimit || pkg.max === "" ? null : pkg.max,
      noLimit: Boolean(pkg.noLimit),
      planFee: pkg.planFee === "" ? null : pkg.planFee,
    })),
  };
}

function getLegacyMigrationSource() {
  const versions = parseJsonStorage(LEGACY_VERSIONS_KEY, []);
  const legacyConfig = parseJsonStorage(LEGACY_CONFIG_KEY, null);
  const drafts = parseJsonStorage(LEGACY_DRAFTS_KEY, []);

  let models = Array.isArray(versions)
    ? versions.map((item, index) => normalizeLegacyModel(item, `Legacy Allocation ${index + 1}`)).filter(Boolean)
    : [];

  if (models.length === 0 && legacyConfig) {
    const normalized = normalizeLegacyModel(legacyConfig, "Legacy Current Allocation");
    if (normalized) models = [{ ...normalized, status: "Active" }];
  }

  const normalizedDrafts = Array.isArray(drafts)
    ? drafts.map((item, index) => normalizeLegacyModel({ ...item, status: "Draft" }, `Legacy Draft ${index + 1}`)).filter(Boolean)
    : [];

  return { models, drafts: normalizedDrafts };
}

function collapseLegacyActiveConflicts(models) {
  const groups = new Map();
  models.forEach((model) => {
    const key = `${normalizeValue(model.chainName).toLowerCase()}::${normalizeValue(model.allocationPurpose).toLowerCase()}`;
    const rows = groups.get(key) || [];
    rows.push(model);
    groups.set(key, rows);
  });

  const result = [];
  groups.forEach((rows) => {
    const activeRows = rows
      .filter((model) => normalizeValue(model.status).toLowerCase() === "active")
      .sort((left, right) =>
        normalizeValue(right.effectiveFrom).localeCompare(normalizeValue(left.effectiveFrom)),
      );
    const retainedActive = activeRows[0] || null;

    rows.forEach((model) => {
      const normalizedStatus = normalizeValue(model.status).toLowerCase();
      if (normalizedStatus === "draft") return;
      result.push({
        ...model,
        status:
          retainedActive && model === retainedActive
            ? "Active"
            : "Inactive",
      });
    });
  });

  return result;
}

export async function ensureAllocationBackendMigration() {
  const existing = await billingApiRequest("/settings/allocations");
  if (Array.isArray(existing) && existing.length > 0) {
    if (typeof window !== "undefined") {
      window.localStorage.setItem(MIGRATION_MARKER_KEY, "true");
    }
    return { migrated: false, records: existing, skipped: [] };
  }

  const { models, drafts } = getLegacyMigrationSource();
  if (models.length === 0 && drafts.length === 0) {
    return { migrated: false, records: [], skipped: [] };
  }

  const skipped = [];
  const normalizedModels = collapseLegacyActiveConflicts(models);
  const inactive = normalizedModels
    .filter((model) => model.status === "Inactive")
    .sort((left, right) =>
      normalizeValue(left.effectiveFrom).localeCompare(normalizeValue(right.effectiveFrom)),
    );
  const active = normalizedModels
    .filter((model) => model.status === "Active")
    .sort((left, right) =>
      normalizeValue(left.effectiveFrom).localeCompare(normalizeValue(right.effectiveFrom)),
    );

  for (const model of inactive) {
    if (!isBackendActivatable(model)) {
      skipped.push(model.name || model.chainName || "Legacy inactive Allocation");
      await billingApiRequest("/settings/allocations/drafts", {
        method: "POST",
        body: JSON.stringify(modelToBackendPayload(model)),
      });
      continue;
    }

    const created = await billingApiRequest("/settings/allocations/activate", {
      method: "POST",
      body: JSON.stringify(modelToBackendPayload(model)),
    });
    await billingApiRequest(`/settings/allocations/${created.id || created.modelId}/inactivate`, {
      method: "POST",
      body: JSON.stringify({
        effectiveTo: /^\d{4}-\d{2}-\d{2}$/.test(model.effectiveTo)
          ? model.effectiveTo
          : model.effectiveFrom,
        remarks: "Migrated from browser Allocation history",
      }),
    });
  }

  for (const model of active) {
    if (!isBackendActivatable(model)) {
      skipped.push(model.name || model.chainName || "Legacy active Allocation");
      await billingApiRequest("/settings/allocations/drafts", {
        method: "POST",
        body: JSON.stringify(modelToBackendPayload(model)),
      });
      continue;
    }

    await billingApiRequest("/settings/allocations/activate", {
      method: "POST",
      body: JSON.stringify(modelToBackendPayload(model)),
    });
  }

  for (const draft of drafts) {
    await billingApiRequest("/settings/allocations/drafts", {
      method: "POST",
      body: JSON.stringify(modelToBackendPayload(draft)),
    });
  }

  const records = await billingApiRequest("/settings/allocations");
  if (typeof window !== "undefined") {
    window.localStorage.setItem(MIGRATION_MARKER_KEY, "true");
  }
  return { migrated: true, records: Array.isArray(records) ? records : [], skipped };
}
