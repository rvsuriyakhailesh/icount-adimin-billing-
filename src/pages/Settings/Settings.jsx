import { useEffect, useMemo, useState } from "react";
import { billingApiRequest, ensureAllocationBackendMigration } from "../../utils/billingApi";
import "./Settings.css";

function normalizeValue(value) {
  return String(value || "").trim();
}

function getAllocationIdentity(model, fallback = "") {
  const explicitPurpose =
    normalizeValue(model?.allocationPurpose) ||
    normalizeValue(model?.purpose);

  const rawIdentity =
    normalizeValue(model?.chainName) ||
    normalizeValue(model?.networkName) ||
    normalizeValue(model?.name) ||
    normalizeValue(model?.modelName) ||
    normalizeValue(fallback);

  let chainName = rawIdentity;
  let allocationPurpose = explicitPurpose;

  // Backward compatibility for records previously saved as
  // "UYLJK Stage 3" or "UYLJK - Stage 3".
  // The Chain identity must remain UYLJK even when Purpose is already
  // stored separately as Stage 3.
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

  const displayName = [chainName, allocationPurpose].filter(Boolean).join(" ");

  return {
    chainName,
    allocationPurpose,
    displayName: displayName || rawIdentity || normalizeValue(fallback),
  };
}

function getAllocationChainName(model, fallback = "") {
  return getAllocationIdentity(model, fallback).chainName;
}

function getAllocationPurpose(model, fallback = "") {
  return (
    getAllocationIdentity(model, fallback).allocationPurpose ||
    normalizeValue(fallback)
  );
}

function getAllocationDisplayName(model, fallback = "") {
  return getAllocationIdentity(model, fallback).displayName;
}


function formatDateInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function addDays(dateValue, days) {
  const parsed = new Date(`${dateValue}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) return "";
  parsed.setDate(parsed.getDate() + days);
  return formatDateInput(parsed);
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
    status:
      normalizeValue(model.status) === "Inactive" ||
      normalizeValue(model.status) === "Previous"
        ? "Inactive"
        : normalizeValue(model.status) === "Active"
          ? "Active"
          : "",
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
    createdAt: normalizeValue(model.createdAt),
    activatedAt: normalizeValue(model.activatedAt),
    sourceModelId: normalizeValue(model.sourceModelId),
    activationReason: normalizeValue(model.activationReason),
  };
}

function normalizeLoadedAllocationModels(models) {
  const normalized = models
    .map((model, index) =>
      normalizeAllocationModel(model, `Chain Allocation ${index + 1}`),
    )
    .filter(Boolean);

  if (normalized.length === 0) return [];

  const grouped = new Map();
  normalized.forEach((model) => {
    const chainKey = getAllocationChainName(model, "Unassigned").toLowerCase();
    const purposeKey =
      getAllocationPurpose(model).toLowerCase() || "default";
    const key = `${chainKey}::${purposeKey}`;
    const rows = grouped.get(key) || [];
    rows.push(model);
    grouped.set(key, rows);
  });

  const result = [];
  grouped.forEach((rows) => {
    const activeCandidates = rows
      .filter((model) => model.status === "Active")
      .sort((left, right) =>
        normalizeValue(right.effectiveFrom).localeCompare(
          normalizeValue(left.effectiveFrom),
        ),
      );

    const activeModel =
      activeCandidates[0] ||
      [...rows].sort((left, right) =>
        normalizeValue(right.effectiveFrom).localeCompare(
          normalizeValue(left.effectiveFrom),
        ),
      )[0];

    rows.forEach((model) => {
      result.push({
        ...model,
        chainName: getAllocationChainName(model, "Unassigned"),
        allocationPurpose: getAllocationPurpose(model),
        name: getAllocationDisplayName(model, "Unassigned"),
        status:
          model.status === "Inactive" || model.status === "Previous"
            ? "Inactive"
            : model.modelId === activeModel?.modelId
              ? "Active"
              : "Inactive",
        effectiveTo:
          model.status === "Active" && model.modelId === activeModel?.modelId
            ? ""
            : normalizeValue(model.effectiveTo),
      });
    });
  });

  return result;
}

function formatAllocationModelStructure(model) {
  if (!model || !Array.isArray(model.slabs) || model.slabs.length === 0) {
    return "-";
  }

  return model.slabs
    .map((slab) => `${slab.label} → ₹${normalizeValue(slab.planFee) || "0"}`)
    .join(" | ");
}

function normalizeMappingRow(mapping, field) {
  return {
    mappingId: mapping.mappingId,
    product: normalizeValue(mapping.product),
    fieldId: field.fieldId,
    billingIncField: normalizeValue(field.billingIncField),
    erpHeader: normalizeValue(field.erpHeader),
    status: field.status === "Inactive" ? "Inactive" : "Active",
  };
}

function buildRowsFromMappings(erpMappings = []) {
  return erpMappings.flatMap((mapping) => {
    const rows = [];

    // Backward compatibility for earlier fixed fields.
    if (normalizeValue(mapping.glAccount)) {
      rows.push({
        fieldId: `${mapping.mappingId}-gl-account`,
        billingIncField: "GL Account",
        erpHeader: "GL Account",
        status: mapping.status === "Inactive" ? "Inactive" : "Active",
      });
    }

    if (normalizeValue(mapping.costCenter)) {
      rows.push({
        fieldId: `${mapping.mappingId}-cost-center`,
        billingIncField: "Cost Center",
        erpHeader: "Cost Center",
        status: mapping.status === "Inactive" ? "Inactive" : "Active",
      });
    }

    const additionalFields = Array.isArray(mapping.additionalFields)
      ? mapping.additionalFields
      : [];

    additionalFields.forEach((field) => {
      const normalized = {
        fieldId:
          field.fieldId ||
          `erp-field-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        billingIncField: normalizeValue(field.billingIncField),
        erpHeader: normalizeValue(field.erpHeader),
        status: field.status === "Inactive" ? "Inactive" : "Active",
      };

      const duplicate = rows.some(
        (row) =>
          row.billingIncField.toLowerCase() ===
            normalized.billingIncField.toLowerCase() &&
          row.erpHeader.toLowerCase() === normalized.erpHeader.toLowerCase(),
      );

      if (!duplicate) {
        rows.push(normalized);
      }
    });

    return rows.map((field) => normalizeMappingRow(mapping, field));
  });
}

function Settings({
  erpMappings = [],
  setErpMappings = () => {},
}) {
  const [product, setProduct] = useState("Billing INC");
  const [billingIncField, setBillingIncField] = useState("");
  const [erpHeader, setErpHeader] = useState("");
  const [editingRowKey, setEditingRowKey] = useState("");
  const [searchTerm, setSearchTerm] = useState("");
  const [saveMessage, setSaveMessage] = useState("");
  const [settingsView, setSettingsView] = useState("erp");
  const [allocationModels, setAllocationModels] = useState([]);
  const [selectedAllocationChain, setSelectedAllocationChain] = useState("");
  const [allocationDrafts, setAllocationDrafts] = useState([]);
  const [allocationMessage, setAllocationMessage] = useState("");
  const [allocationBackendLoading, setAllocationBackendLoading] = useState(true);
  const [allocationWorkspace, setAllocationWorkspace] = useState("current");
  const [allocationEditorMode, setAllocationEditorMode] = useState("new");
  const [newAllocationModel, setNewAllocationModel] = useState({
    draftId: "",
    sourceModelId: "",
    chainName: "",
    allocationPurpose: "",
    name: "",
    effectiveFrom: "",
    otfPricePerDevice: "",
    activationReason: "",
    packages: [
      { packageId: `package-${Date.now()}-1`, packageName: "", min: "", max: "", noLimit: false, planFee: "" },
    ],
  });

  function allocationDraftToApiPayload(draft) {
    const chainName = getAllocationChainName(draft);
    const allocationPurpose = getAllocationPurpose(draft);
    const displayName = [chainName, allocationPurpose].filter(Boolean).join(" ");

    return {
      chainName: chainName || null,
      allocationPurpose: allocationPurpose || null,
      name: displayName || normalizeValue(draft?.name) || null,
      displayName: displayName || normalizeValue(draft?.name) || null,
      effectiveFrom: normalizeValue(draft?.effectiveFrom) || null,
      otfPricePerDevice:
        normalizeValue(draft?.otfPricePerDevice) === ""
          ? null
          : normalizeValue(draft?.otfPricePerDevice),
      sourceModelId: normalizeValue(draft?.sourceModelId) || null,
      activationReason: normalizeValue(draft?.activationReason) || null,
      remarks: normalizeValue(draft?.remarks) || null,
      packages: (Array.isArray(draft?.packages) ? draft.packages : []).map((item) => ({
        packageId: normalizeValue(item?.packageId) || null,
        packageName: normalizeValue(item?.packageName) || null,
        min: normalizeValue(item?.min) === "" ? null : item.min,
        max:
          item?.noLimit || normalizeValue(item?.max) === ""
            ? null
            : item.max,
        noLimit: Boolean(item?.noLimit),
        planFee: normalizeValue(item?.planFee) === "" ? null : item.planFee,
      })),
    };
  }

  async function refreshBackendAllocations({ preserveSelection = true } = {}) {
    const records = await billingApiRequest("/settings/allocations");
    const rows = Array.isArray(records) ? records : [];
    const models = rows.filter((model) => model.status !== "Draft");
    const drafts = rows.filter((model) => model.status === "Draft");

    setAllocationModels(models);
    setAllocationDrafts(drafts);

    const activeModels = models.filter((model) => model.status === "Active");
    const currentSelectionStillExists = activeModels.some(
      (model) => normalizeValue(model.modelId) === normalizeValue(selectedAllocationChain),
    );

    if (!preserveSelection || !currentSelectionStillExists) {
      setSelectedAllocationChain(activeModels[0]?.modelId || "");
    }

    window.dispatchEvent(
      new CustomEvent("billing-allocation-settings-updated", {
        detail: rows,
      }),
    );

    return rows;
  }

  useEffect(() => {
    let cancelled = false;

    async function initializeAllocationBackend() {
      setAllocationBackendLoading(true);
      try {
        const migration = await ensureAllocationBackendMigration();
        if (cancelled) return;

        const rows = Array.isArray(migration?.records)
          ? migration.records
          : await billingApiRequest("/settings/allocations");
        if (cancelled) return;

        const models = (Array.isArray(rows) ? rows : []).filter(
          (model) => model.status !== "Draft",
        );
        const drafts = (Array.isArray(rows) ? rows : []).filter(
          (model) => model.status === "Draft",
        );
        const active = models.find((model) => model.status === "Active");

        setAllocationModels(models);
        setAllocationDrafts(drafts);
        setSelectedAllocationChain(active?.modelId || "");

        if (migration?.migrated) {
          const skippedCount = Array.isArray(migration.skipped)
            ? migration.skipped.length
            : 0;
          setAllocationMessage(
            skippedCount > 0
              ? `Browser Allocation data migrated to backend. ${skippedCount} incomplete historical Allocation(s) were preserved as Drafts.`
              : "Browser Allocation data migrated to backend successfully. PostgreSQL is now the Allocation source of truth.",
          );
        }
      } catch (error) {
        if (!cancelled) {
          console.error("Unable to initialize Allocation backend", error);
          setAllocationMessage(
            `Unable to load Allocation data from backend. ${error.message}`,
          );
        }
      } finally {
        if (!cancelled) setAllocationBackendLoading(false);
      }
    }

    initializeAllocationBackend();

    return () => {
      cancelled = true;
    };
  }, []);

  const savedRows = useMemo(() => {
    const rows = buildRowsFromMappings(erpMappings);

    return rows.sort((left, right) => {
      if (left.status !== right.status) {
        return left.status === "Active" ? -1 : 1;
      }

      const productCompare = left.product.localeCompare(right.product);
      if (productCompare !== 0) {
        return productCompare;
      }

      return left.billingIncField.localeCompare(right.billingIncField);
    });
  }, [erpMappings]);

  const filteredRows = useMemo(() => {
    const normalizedSearch = normalizeValue(searchTerm).toLowerCase();

    if (!normalizedSearch) {
      return savedRows;
    }

    return savedRows.filter((row) =>
      [
        row.product,
        row.billingIncField,
        row.erpHeader,
        row.status,
      ]
        .map((value) => normalizeValue(value).toLowerCase())
        .some((value) => value.includes(normalizedSearch)),
    );
  }, [savedRows, searchTerm]);

  function resetEditor() {
    setProduct("Billing INC");
    setBillingIncField("");
    setErpHeader("");
    setEditingRowKey("");
    setSaveMessage("");
  }

  function handleSave() {
    const nextProduct = normalizeValue(product);
    const nextBillingIncField = normalizeValue(billingIncField);
    const nextErpHeader = normalizeValue(erpHeader);

    if (!nextProduct) {
      setSaveMessage("Product is required.");
      return;
    }

    if (!nextBillingIncField) {
      setSaveMessage("Billing INC Field is required.");
      return;
    }

    if (!nextErpHeader) {
      setSaveMessage("ERP Header is required.");
      return;
    }

    const editingRow = savedRows.find(
      (row) =>
        `${row.mappingId}::${row.fieldId}` === editingRowKey,
    );

    if (editingRow) {
      setErpMappings((currentMappings) =>
        currentMappings.map((mapping) => {
          if (mapping.mappingId !== editingRow.mappingId) {
            return mapping;
          }

          const nextAdditionalFields = Array.isArray(mapping.additionalFields)
            ? mapping.additionalFields.map((field) =>
                field.fieldId === editingRow.fieldId
                  ? {
                      ...field,
                      billingIncField: nextBillingIncField,
                      erpHeader: nextErpHeader,
                    }
                  : field,
              )
            : [];

          const legacyGlRow =
            editingRow.fieldId === `${mapping.mappingId}-gl-account`;
          const legacyCostCenterRow =
            editingRow.fieldId === `${mapping.mappingId}-cost-center`;

          if (legacyGlRow || legacyCostCenterRow) {
            nextAdditionalFields.push({
              fieldId: editingRow.fieldId,
              billingIncField: nextBillingIncField,
              erpHeader: nextErpHeader,
              status: editingRow.status,
            });
          }

          return {
            ...mapping,
            product: nextProduct,
            glAccount: legacyGlRow ? "" : mapping.glAccount,
            costCenter: legacyCostCenterRow ? "" : mapping.costCenter,
            additionalFields: nextAdditionalFields,
            updatedAt: new Date().toISOString(),
          };
        }),
      );

      setSaveMessage("ERP mapping updated.");
      setEditingRowKey("");
      setBillingIncField("");
      setErpHeader("");
      return;
    }

    const duplicate = savedRows.some(
      (row) =>
        row.product.toLowerCase() === nextProduct.toLowerCase() &&
        row.billingIncField.toLowerCase() ===
          nextBillingIncField.toLowerCase() &&
        row.erpHeader.toLowerCase() === nextErpHeader.toLowerCase(),
    );

    if (duplicate) {
      setSaveMessage("This ERP mapping already exists.");
      return;
    }

    setErpMappings((currentMappings) => {
      const existingProductIndex = currentMappings.findIndex(
        (mapping) =>
          normalizeValue(mapping.product).toLowerCase() ===
          nextProduct.toLowerCase(),
      );

      const newField = {
        fieldId: `erp-field-${Date.now()}-${Math.random()
          .toString(36)
          .slice(2, 7)}`,
        billingIncField: nextBillingIncField,
        erpHeader: nextErpHeader,
        status: "Active",
      };

      if (existingProductIndex >= 0) {
        const nextMappings = [...currentMappings];
        const currentMapping = nextMappings[existingProductIndex];

        nextMappings[existingProductIndex] = {
          ...currentMapping,
          product: nextProduct,
          additionalFields: [
            ...(Array.isArray(currentMapping.additionalFields)
              ? currentMapping.additionalFields
              : []),
            newField,
          ],
          updatedAt: new Date().toISOString(),
        };

        return nextMappings;
      }

      return [
        ...currentMappings,
        {
          mappingId: `erp-mapping-${Date.now()}`,
          product: nextProduct,
          status: "Active",
          glAccount: "",
          costCenter: "",
          additionalFields: [newField],
          updatedAt: new Date().toISOString(),
        },
      ];
    });

    setBillingIncField("");
    setErpHeader("");
    setSaveMessage("ERP mapping saved.");
  }

  function handleEdit(row) {
    setProduct(row.product);
    setBillingIncField(row.billingIncField);
    setErpHeader(row.erpHeader);
    setEditingRowKey(`${row.mappingId}::${row.fieldId}`);
    setSaveMessage("");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function handleToggleStatus(row) {
    setErpMappings((currentMappings) =>
      currentMappings.map((mapping) => {
        if (mapping.mappingId !== row.mappingId) {
          return mapping;
        }

        const nextStatus = row.status === "Active" ? "Inactive" : "Active";
        const additionalFields = Array.isArray(mapping.additionalFields)
          ? mapping.additionalFields.map((field) =>
              field.fieldId === row.fieldId
                ? {
                    ...field,
                    status: nextStatus,
                  }
                : field,
            )
          : [];

        const legacyGlRow =
          row.fieldId === `${mapping.mappingId}-gl-account`;
        const legacyCostCenterRow =
          row.fieldId === `${mapping.mappingId}-cost-center`;

        if (legacyGlRow || legacyCostCenterRow) {
          additionalFields.push({
            fieldId: row.fieldId,
            billingIncField: row.billingIncField,
            erpHeader: row.erpHeader,
            status: nextStatus,
          });
        }

        return {
          ...mapping,
          glAccount: legacyGlRow ? "" : mapping.glAccount,
          costCenter: legacyCostCenterRow ? "" : mapping.costCenter,
          additionalFields,
          updatedAt: new Date().toISOString(),
        };
      }),
    );

    if (editingRowKey === `${row.mappingId}::${row.fieldId}`) {
      resetEditor();
    }
  }

  const activeAllocationModels = useMemo(
    () =>
      allocationModels
        .filter((model) => model.status === "Active")
        .sort((left, right) =>
          getAllocationDisplayName(left).localeCompare(
            getAllocationDisplayName(right),
          ),
        ),
    [allocationModels],
  );

  const currentAllocationModel = useMemo(() => {
    if (activeAllocationModels.length === 0) return null;

    const selected = activeAllocationModels.find(
      (model) =>
        normalizeValue(model.modelId) ===
        normalizeValue(selectedAllocationChain),
    );

    // Backward compatibility if an older in-memory selection still
    // contains only the Chain Name.
    const legacySelected = activeAllocationModels.find(
      (model) =>
        getAllocationChainName(model).toLowerCase() ===
        normalizeValue(selectedAllocationChain).toLowerCase(),
    );

    return selected || legacySelected || activeAllocationModels[0];
  }, [activeAllocationModels, selectedAllocationChain]);

  const inactiveAllocationModels = useMemo(
    () =>
      allocationModels
        .filter((model) => model.status !== "Active")
        .sort((left, right) => {
          const chainCompare = getAllocationChainName(left).localeCompare(
            getAllocationChainName(right),
          );
          if (chainCompare !== 0) return chainCompare;
          return normalizeValue(right.effectiveFrom).localeCompare(
            normalizeValue(left.effectiveFrom),
          );
        }),
    [allocationModels],
  );

  function createBlankAllocationDraft() {
    return {
      draftId: "",
      sourceModelId: "",
      chainName: "",
      allocationPurpose: "",
      name: "",
      effectiveFrom: "",
      otfPricePerDevice: "",
      activationReason: "",
      packages: [
        {
          packageId: `package-${Date.now()}-1`,
          packageName: "",
          min: "",
          max: "",
          noLimit: false,
          planFee: "",
        },
      ],
    };
  }

  function createAllocationDraftFromModel(model, options = {}) {
    const sourcePackages = Array.isArray(model?.slabs) ? model.slabs : [];

    return {
      draftId: normalizeValue(options.draftId),
      sourceModelId: normalizeValue(model?.modelId),
      chainName: getAllocationChainName(model),
      allocationPurpose: getAllocationPurpose(model),
      name: getAllocationDisplayName(model),
      effectiveFrom: normalizeValue(options.effectiveFrom),
      otfPricePerDevice: normalizeValue(model?.otfPricePerDevice),
      activationReason: normalizeValue(options.activationReason),
      packages:
        sourcePackages.length > 0
          ? sourcePackages.map((slab, index) => ({
              packageId: `package-${Date.now()}-${index + 1}`,
              packageName: normalizeValue(slab?.label),
              min: slab?.min === null || slab?.min === undefined ? "" : String(slab.min),
              max: slab?.max === null || slab?.max === undefined ? "" : String(slab.max),
              noLimit: slab?.max === null || slab?.max === undefined,
              planFee: normalizeValue(slab?.planFee),
            }))
          : createBlankAllocationDraft().packages,
    };
  }

  function createEditorFromSavedDraft(draft) {
    return {
      draftId: normalizeValue(draft?.draftId),
      sourceModelId: normalizeValue(draft?.sourceModelId),
      chainName: getAllocationChainName(draft),
      allocationPurpose: getAllocationPurpose(draft),
      name: getAllocationDisplayName(draft),
      effectiveFrom: normalizeValue(draft?.effectiveFrom),
      otfPricePerDevice: normalizeValue(draft?.otfPricePerDevice),
      activationReason: normalizeValue(draft?.activationReason),
      packages: Array.isArray(draft?.packages) && draft.packages.length > 0
        ? draft.packages.map((item, index) => ({
            packageId: normalizeValue(item?.packageId) || `package-${Date.now()}-${index + 1}`,
            packageName: normalizeValue(item?.packageName),
            min: normalizeValue(item?.min),
            max: normalizeValue(item?.max),
            noLimit: Boolean(item?.noLimit),
            planFee: normalizeValue(item?.planFee),
          }))
        : createBlankAllocationDraft().packages,
    };
  }

  function openAllocationWorkspace(workspace, message = "") {
    setAllocationMessage(message);
    setAllocationWorkspace(workspace);
    window.setTimeout(() => window.scrollTo({ top: 0, behavior: "smooth" }), 0);
  }

  function handleCreateNewAllocation() {
    setAllocationEditorMode("new");
    setNewAllocationModel(createBlankAllocationDraft());
    openAllocationWorkspace(
      "new",
      "New Allocation opened. Enter the actual Chain Name (for example UYLJK), the Allocation Purpose when required (for example Stage 3), package names, quantity ranges and Plan Fees. Only one Active Allocation is allowed for the same Chain and Purpose.",
    );
  }

  async function handleViewAllocationDrafts() {
    try {
      await refreshBackendAllocations();
      openAllocationWorkspace("drafts");
    } catch (error) {
      setAllocationMessage(`Unable to load Allocation Drafts. ${error.message}`);
    }
  }

  async function handleViewInactiveAllocations() {
    try {
      await refreshBackendAllocations();
      openAllocationWorkspace("history");
    } catch (error) {
      setAllocationMessage(`Unable to load Inactive Allocations. ${error.message}`);
    }
  }

  function handleBackToCurrentAllocation() {
    setAllocationMessage("");
    setAllocationWorkspace("current");
  }

  function handleUseInactiveAllocation(model) {
    if (!model) return;
    setAllocationEditorMode("modify");
    setNewAllocationModel(createAllocationDraftFromModel(model));
    openAllocationWorkspace(
      "new",
      `${normalizeValue(model.name) || "Selected Allocation Model"} copied as an editable New Allocation.`,
    );
  }

  function handlePrepareInactiveAllocationActivation(model) {
    if (!model) return;
    setAllocationEditorMode("reactivate");
    setNewAllocationModel(createAllocationDraftFromModel(model));
    openAllocationWorkspace(
      "new",
      `${normalizeValue(model.name) || "Inactive Allocation"} selected for reactivation. Enter a new Effective From Date and Reason, then click Activate Model.`,
    );
  }

  async function handleMoveCurrentAllocationToInactive() {
    if (!currentAllocationModel) {
      setAllocationMessage("No Active Allocation is available for the selected Chain Network.");
      return;
    }

    const chainName = getAllocationChainName(currentAllocationModel);
    const confirmed = window.confirm(
      `Move ${getAllocationDisplayName(currentAllocationModel) || chainName} from Active Allocation to Inactive Allocations?`,
    );
    if (!confirmed) return;

    try {
      const today = formatDateInput(new Date());
      await billingApiRequest(
        `/settings/allocations/${currentAllocationModel.modelId}/inactivate`,
        {
          method: "POST",
          body: JSON.stringify({
            effectiveTo: today,
            remarks: "Moved to Inactive manually from Settings",
          }),
        },
      );
      await refreshBackendAllocations({ preserveSelection: false });
      setAllocationMessage(`${getAllocationDisplayName(currentAllocationModel) || chainName} moved to Inactive Allocations manually.`);
    } catch (error) {
      setAllocationMessage(`Unable to move Allocation to Inactive. ${error.message}`);
    }
  }

  function handleAllocationPackageChange(packageId, field, value) {
    if (allocationEditorMode === "reactivate") return;
    setAllocationMessage("");
    setNewAllocationModel((current) => ({
      ...current,
      packages: current.packages.map((item) => {
        if (item.packageId !== packageId) return item;

        if (field === "noLimit") {
          return {
            ...item,
            noLimit: Boolean(value),
            max: value ? "" : item.max,
          };
        }

        if (["min", "max"].includes(field)) {
          return { ...item, [field]: String(value || "").replace(/\D/g, "") };
        }

        if (field === "planFee") {
          return {
            ...item,
            planFee: String(value || "").replace(/[^0-9.]/g, ""),
          };
        }

        return { ...item, [field]: value };
      }),
    }));
  }

  function handleAddAllocationPackage() {
    if (allocationEditorMode === "reactivate") return;
    setAllocationMessage("");
    setNewAllocationModel((current) => ({
      ...current,
      packages: [
        ...current.packages,
        {
          packageId: `package-${Date.now()}-${current.packages.length + 1}`,
          packageName: "",
          min: "",
          max: "",
          noLimit: false,
          planFee: "",
        },
      ],
    }));
  }

  function handleRemoveAllocationPackage(packageId) {
    if (allocationEditorMode === "reactivate") return;
    setAllocationMessage("");
    setNewAllocationModel((current) => ({
      ...current,
      packages:
        current.packages.length <= 1
          ? current.packages
          : current.packages.filter((item) => item.packageId !== packageId),
    }));
  }

  function validateAllocationForActivation(draft) {
    const chainName = getAllocationChainName(draft);
    const allocationPurpose = getAllocationPurpose(draft);
    const displayName = [chainName, allocationPurpose].filter(Boolean).join(" ");
    const effectiveFrom = normalizeValue(draft.effectiveFrom);
    const packages = Array.isArray(draft.packages) ? draft.packages : [];

    if (!chainName) return { error: "Chain Name is required." };
    if (!/^\d{4}-\d{2}-\d{2}$/.test(effectiveFrom)) {
      return { error: "Effective From Date is required before activation." };
    }
    if (packages.length === 0) return { error: "Add at least one Plan Package." };

    const normalizedPackages = [];
    for (let index = 0; index < packages.length; index += 1) {
      const item = packages[index];
      const packageName = normalizeValue(item.packageName);
      const min = Number(item.min);
      const max = item.noLimit ? null : Number(item.max);
      const planFee = normalizeValue(item.planFee);

      if (!packageName) return { error: `Enter Package Name for row ${index + 1}.` };
      if (!Number.isInteger(min) || min < 1) {
        return { error: `Enter a valid From Quantity for ${packageName}.` };
      }
      if (!item.noLimit && (!Number.isInteger(max) || max < min)) {
        return { error: `To Quantity for ${packageName} must be greater than or equal to From Quantity.` };
      }
      if (planFee === "" || !Number.isFinite(Number(planFee)) || Number(planFee) < 0) {
        return { error: `Enter a valid Plan Fee for ${packageName}.` };
      }

      normalizedPackages.push({
        slabId: `slab-${Date.now()}-${index + 1}`,
        label: packageName,
        min,
        max,
        planFee,
      });
    }

    const duplicatePackageName = normalizedPackages.some((item, index) =>
      normalizedPackages.some(
        (other, otherIndex) =>
          index !== otherIndex &&
          item.label.toLowerCase() === other.label.toLowerCase(),
      ),
    );
    if (duplicatePackageName) {
      return { error: "Plan Package names must be unique within the model." };
    }

    const orderedPackages = [...normalizedPackages].sort((a, b) => a.min - b.min);
    for (let index = 1; index < orderedPackages.length; index += 1) {
      const previous = orderedPackages[index - 1];
      const current = orderedPackages[index];
      if (previous.max === null || current.min <= previous.max) {
        return { error: `Plan Package ranges overlap between ${previous.label} and ${current.label}.` };
      }
    }

    if (orderedPackages.filter((item) => item.max === null).length > 1) {
      return { error: "Only one Plan Package can use No Limit." };
    }

    const normalizedPurpose = allocationPurpose.toLowerCase();
    const sameChainCurrentModel = allocationModels.find(
      (model) =>
        model.status === "Active" &&
        getAllocationChainName(model).toLowerCase() === chainName.toLowerCase() &&
        getAllocationPurpose(model).toLowerCase() === normalizedPurpose,
    );

    if (
      sameChainCurrentModel?.effectiveFrom &&
      effectiveFrom <= sameChainCurrentModel.effectiveFrom
    ) {
      return {
        error: `Effective From must be after the current ${chainName} Allocation Effective From (${sameChainCurrentModel.effectiveFrom}).`,
      };
    }

    return {
      orderedPackages,
      chainName,
      allocationPurpose,
      displayName,
      effectiveFrom,
      sameChainCurrentModel,
    };
  }

  async function handleSaveAllocationDraft() {
    const draftChainName = getAllocationChainName(newAllocationModel);
    const draftPurpose = getAllocationPurpose(newAllocationModel);
    const draftDisplayName =
      [draftChainName, draftPurpose].filter(Boolean).join(" ") ||
      "Untitled Chain Allocation Draft";
    const payload = allocationDraftToApiPayload({
      ...newAllocationModel,
      chainName: draftChainName,
      allocationPurpose: draftPurpose,
      name: draftDisplayName,
    });

    try {
      const existingDraftId = normalizeValue(newAllocationModel.draftId);
      const saved = existingDraftId
        ? await billingApiRequest(`/settings/allocations/${existingDraftId}`, {
            method: "PATCH",
            body: JSON.stringify(payload),
          })
        : await billingApiRequest("/settings/allocations/drafts", {
            method: "POST",
            body: JSON.stringify(payload),
          });

      await refreshBackendAllocations();
      setNewAllocationModel(createEditorFromSavedDraft(saved));
      setAllocationEditorMode("draft");
      openAllocationWorkspace(
        "drafts",
        `${saved.name || draftDisplayName} saved as Draft in backend. Current Allocation is unchanged.`,
      );
    } catch (error) {
      setAllocationMessage(`Unable to save Allocation Draft. ${error.message}`);
    }
  }

  async function activateAllocation(draft) {
    const validation = validateAllocationForActivation(draft);
    if (validation.error) {
      setAllocationMessage(validation.error);
      return false;
    }

    if (validation.sameChainCurrentModel) {
      setAllocationMessage(
        `${validation.chainName} already has an Active Allocation for ${validation.allocationPurpose || "the default purpose"}. Move the current Allocation to Inactive Allocations manually before activating the new Allocation.`,
      );
      return false;
    }

    const confirmed = window.confirm(
      `Activate ${validation.displayName || validation.chainName} Allocation from ${validation.effectiveFrom}?`,
    );
    if (!confirmed) return false;

    try {
      const payload = allocationDraftToApiPayload({
        ...draft,
        chainName: validation.chainName,
        allocationPurpose: validation.allocationPurpose,
        name: validation.displayName || validation.chainName,
        packages: validation.orderedPackages.map((item, index) => ({
          packageId: item.slabId || `package-${index + 1}`,
          packageName: item.label,
          min: item.min,
          max: item.max === null ? "" : item.max,
          noLimit: item.max === null,
          planFee: item.planFee,
        })),
      });

      const existingDraftId = normalizeValue(draft?.draftId);

      const activated = existingDraftId
        ? await billingApiRequest(
            `/settings/allocations/${existingDraftId}/activate`,
            {
              method: "POST",
              body: JSON.stringify({
                activationReason: normalizeValue(draft?.activationReason) || null,
                remarks: normalizeValue(draft?.remarks) || null,
              }),
            },
          )
        : await billingApiRequest("/settings/allocations/activate", {
            method: "POST",
            body: JSON.stringify(payload),
          });

      await refreshBackendAllocations({ preserveSelection: false });
      setSelectedAllocationChain(activated.modelId || activated.id || "");
      setNewAllocationModel(createBlankAllocationDraft());
      setAllocationEditorMode("new");
      setAllocationWorkspace("current");
      setAllocationMessage(
        `${activated.name || validation.displayName || validation.chainName} is now the Active Allocation from ${activated.effectiveFrom || validation.effectiveFrom}.`,
      );
      return true;
    } catch (error) {
      setAllocationMessage(`Unable to activate Allocation. ${error.message}`);
      return false;
    }
  }

  function handleActivateAllocationModel() {
    activateAllocation(newAllocationModel);
  }

  function handleEditAllocationDraft(draft) {
    setAllocationEditorMode("draft");
    setNewAllocationModel(createEditorFromSavedDraft(draft));
    openAllocationWorkspace("new", "Draft opened for editing.");
  }

  function handleActivateSavedDraft(draft) {
    const editorDraft = createEditorFromSavedDraft(draft);
    const validation = validateAllocationForActivation(editorDraft);
    if (validation.error) {
      setAllocationEditorMode("draft");
      setNewAllocationModel(editorDraft);
      openAllocationWorkspace("new", `${validation.error} Complete the Draft before activation.`);
      return;
    }
    activateAllocation(editorDraft);
  }

  async function handleDeleteAllocationDraft(draftId) {
    if (!window.confirm("Delete this Allocation Draft?")) return;
    try {
      await billingApiRequest(`/settings/allocations/${draftId}`, {
        method: "DELETE",
      });
      await refreshBackendAllocations();
      setAllocationMessage("Allocation Draft deleted from backend.");
    } catch (error) {
      setAllocationMessage(`Unable to delete Allocation Draft. ${error.message}`);
    }
  }


  return (
    <section className="erp-settings">
      <div className="erp-settings__header">
        <div>
          <p className="erp-settings__eyebrow">Settings · Billing INC</p>
          {settingsView === "allocation" ? (
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "nowrap" }}>
              <h1 style={{ margin: 0 }}>1B Subscription · Allocation</h1>
              <span
                role="img"
                tabIndex={0}
                aria-label="Allocation information"
                title="Maintain Allocation configurations for Chain Networks. If a Chain already has an Active Allocation, move the current Allocation to Inactive before activating a new Allocation for the same Chain."
                style={{
                  width: 24,
                  height: 24,
                  minWidth: 24,
                  border: "1px solid currentColor",
                  borderRadius: "50%",
                  display: "inline-flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: 12,
                  fontWeight: 700,
                  lineHeight: 1,
                  cursor: "help",
                  opacity: 0.78,
                }}
              >
                i
              </span>
            </div>
          ) : (
            <h1>ERP Mapping</h1>
          )}
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <button type="button" className={settingsView === "erp" ? "erp-settings__primary-button" : "erp-settings__secondary-button"} onClick={() => setSettingsView("erp")}>ERP Mapping</button>
          <button type="button" className={settingsView === "allocation" ? "erp-settings__primary-button" : "erp-settings__secondary-button"} onClick={() => {
            setSettingsView("allocation");
            setAllocationWorkspace("current");
            if (!selectedAllocationChain && activeAllocationModels[0]) {
              setSelectedAllocationChain(activeAllocationModels[0].modelId);
            }
          }}>Allocation</button>
        </div>
      </div>

      {settingsView === "allocation" ? (
        <>
          {allocationWorkspace === "current" ? (
            <article className="erp-settings__card">
              <div className="erp-settings__card-heading">
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 8,
                    flexShrink: 0,
                    whiteSpace: "nowrap",
                  }}
                >
                  <h2 style={{ margin: 0, whiteSpace: "nowrap" }}>Current Allocation</h2>
                  <span
                    role="img"
                    tabIndex={0}
                    aria-label="Current Allocation information"
                    title="Each Chain Network can have its own active Allocation. Select the Chain below to view its current package structure."
                    style={{
                      width: 24,
                      height: 24,
                      minWidth: 24,
                      border: "1px solid currentColor",
                      borderRadius: "50%",
                      display: "inline-flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontSize: 12,
                      fontWeight: 700,
                      lineHeight: 1,
                      cursor: "help",
                      opacity: 0.78,
                    }}
                  >
                    i
                  </span>
                </div>
                <div
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    gap: 8,
                    alignItems: "flex-end",
                    justifyContent: "center",
                    marginLeft: "auto",
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      gap: 8,
                      alignItems: "center",
                      justifyContent: "flex-end",
                      flexWrap: "nowrap",
                    }}
                  >
                    <button type="button" className="erp-settings__secondary-button" onClick={handleCreateNewAllocation}>Create New Allocation</button>
                    <button type="button" className="erp-settings__secondary-button" onClick={handleViewAllocationDrafts}>View Drafts ({allocationDrafts.length})</button>
                  </div>

                  <div
                    style={{
                      display: "flex",
                      gap: 8,
                      alignItems: "center",
                      justifyContent: "flex-end",
                      flexWrap: "nowrap",
                    }}
                  >
                    <button type="button" className="erp-settings__secondary-button" onClick={handleMoveCurrentAllocationToInactive} disabled={!currentAllocationModel}>Move to Inactive</button>
                    <button type="button" className="erp-settings__secondary-button" onClick={handleViewInactiveAllocations}>Inactive Allocations</button>
                  </div>
                </div>
              </div>

              {allocationMessage ? <div className="erp-settings__message">{allocationMessage}</div> : null}
              {allocationBackendLoading ? (
                <div className="erp-settings__message">Loading Allocation data from backend...</div>
              ) : null}

              <div className="erp-settings__entry-grid">
                <div className="erp-settings__field">
                  <label>Chain Network</label>
                  <select
                    style={{ minHeight: 42, height: 42, boxSizing: "border-box", fontFamily: "inherit", fontSize: "inherit" }}
                    value={currentAllocationModel ? currentAllocationModel.modelId : ""}
                    onChange={(event) => setSelectedAllocationChain(event.target.value)}
                  >
                    {activeAllocationModels.length === 0 ? (
                      <option value="">No Active Chain Allocation</option>
                    ) : null}
                    {activeAllocationModels.map((model) => (
                      <option key={model.modelId} value={model.modelId}>
                        {getAllocationDisplayName(model)}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="erp-settings__field"><label>Chain Name</label><input type="text" value={getAllocationChainName(currentAllocationModel) || "No Active Allocation"} readOnly /></div>
                <div className="erp-settings__field"><label>Allocation Purpose</label><input type="text" value={getAllocationPurpose(currentAllocationModel) || "-"} readOnly /></div>
                <div className="erp-settings__field"><label>Status</label><input type="text" value={currentAllocationModel ? "Active" : "-"} readOnly /></div>
                <div className="erp-settings__field"><label>Effective From</label><input type="text" value={currentAllocationModel?.effectiveFrom || "-"} readOnly /></div>
                <div className="erp-settings__field"><label>Effective To</label><input type="text" value={currentAllocationModel?.effectiveTo || "Current"} readOnly /></div>
                <div className="erp-settings__field"><label>OTF Price Per Device</label><input type="text" value={currentAllocationModel?.otfPricePerDevice || "-"} readOnly /></div>
              </div>

              <div className="erp-settings__table-wrapper" style={{ marginTop: 16 }}>
                <table className="erp-settings__table">
                  <thead><tr><th>Plan Package</th><th>From Qty</th><th>To Qty</th><th>Plan Fee</th></tr></thead>
                  <tbody>
                    {(currentAllocationModel?.slabs || []).map((slab) => (
                      <tr key={slab.slabId}>
                        <td>{slab.label}</td><td>{slab.min}</td><td>{slab.max === null ? "No Limit" : slab.max}</td><td>₹{Number(slab.planFee || 0).toLocaleString("en-IN")}</td>
                      </tr>
                    ))}
                    {!currentAllocationModel?.slabs?.length ? <tr><td colSpan="4">No Current Allocation Model configured.</td></tr> : null}
                  </tbody>
                </table>
              </div>
            </article>
          ) : null}

          {allocationWorkspace === "new" ? (
            <article className="erp-settings__card" id="new-allocation-model-editor">
              <div className="erp-settings__card-heading">
                <div>
                  <h2>{allocationEditorMode === "reactivate" ? "Reactivate Inactive Allocation" : "New Allocation"}</h2>
                  <p>
                    {allocationEditorMode === "reactivate"
                      ? "The historical package structure is preserved. Enter a new Effective From date and activation reason."
                      : "New Model — Editable. Package names, quantity ranges and Plan Fees are completely user-created."}
                  </p>
                </div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" }}>
                  <button type="button" className="erp-settings__secondary-button" onClick={handleBackToCurrentAllocation}>Back to Current Allocation</button>
                  {allocationEditorMode !== "reactivate" ? <button type="button" className="erp-settings__secondary-button" onClick={handleAddAllocationPackage}>+ Add Plan Package</button> : null}
                </div>
              </div>

              {allocationMessage ? <div className="erp-settings__message">{allocationMessage}</div> : null}

              <div className="erp-settings__entry-grid">
                <div className="erp-settings__field">
                  <label>Chain Name *</label>
                  <input
                    type="text"
                    value={newAllocationModel.chainName || getAllocationChainName(newAllocationModel)}
                    onChange={(event) =>
                      setNewAllocationModel((current) => ({
                        ...current,
                        chainName: event.target.value,
                        name: [
                          event.target.value,
                          normalizeValue(current.allocationPurpose),
                        ].filter(Boolean).join(" "),
                      }))
                    }
                    placeholder="Enter chain name, e.g. UYLJK"
                    readOnly={allocationEditorMode === "reactivate" || Boolean(newAllocationModel.sourceModelId)}
                  />
                </div>
                <div className="erp-settings__field">
                  <label>Allocation Purpose</label>
                  <input
                    type="text"
                    value={newAllocationModel.allocationPurpose || ""}
                    onChange={(event) =>
                      setNewAllocationModel((current) => ({
                        ...current,
                        allocationPurpose: event.target.value,
                        name: [
                          normalizeValue(current.chainName) || getAllocationChainName(current),
                          event.target.value,
                        ].filter(Boolean).join(" "),
                      }))
                    }
                    placeholder="e.g. Stage 3"
                    readOnly={allocationEditorMode === "reactivate"}
                  />
                </div>
                <div className="erp-settings__field"><label>Effective From *</label><input type="date" value={newAllocationModel.effectiveFrom} onChange={(event) => setNewAllocationModel((current) => ({ ...current, effectiveFrom: event.target.value }))} /></div>
                <div className="erp-settings__field"><label>OTF Price Per Device</label><input type="text" inputMode="decimal" value={newAllocationModel.otfPricePerDevice} onChange={(event) => setNewAllocationModel((current) => ({ ...current, otfPricePerDevice: event.target.value.replace(/[^0-9.]/g, "") }))} placeholder={currentAllocationModel?.otfPricePerDevice || "Optional"} readOnly={allocationEditorMode === "reactivate"} /></div>
                <div className="erp-settings__field" style={{ gridColumn: "1 / -1" }}><label>Activation Reason / Remarks</label><input type="text" value={newAllocationModel.activationReason} onChange={(event) => setNewAllocationModel((current) => ({ ...current, activationReason: event.target.value }))} placeholder="Optional for new/modified model; recommended when reactivating a inactive allocation" /></div>
              </div>

              <div className="erp-settings__table-wrapper" style={{ marginTop: 16, width: "100%", maxWidth: "100%", overflowX: "hidden" }}>
                <table className="erp-settings__table" style={{ width: "100%", maxWidth: "100%", tableLayout: "fixed" }}>
                  <colgroup>
                    <col style={{ width: "23%" }} />
                    <col style={{ width: "15%" }} />
                    <col style={{ width: "15%" }} />
                    <col style={{ width: "11%" }} />
                    <col style={{ width: "20%" }} />
                    <col style={{ width: "16%" }} />
                  </colgroup>
                  <thead><tr><th>Plan Package *</th><th>From Qty *</th><th>To Qty *</th><th style={{ textAlign: "center" }}>No Limit</th><th>Plan Fee *</th><th>Action</th></tr></thead>
                  <tbody>
                    {newAllocationModel.packages.map((item) => (
                      <tr key={item.packageId}>
                        <td style={{ minWidth: 0 }}><input style={{ width: "100%", maxWidth: "100%", minWidth: 0, boxSizing: "border-box" }} type="text" value={item.packageName} onChange={(event) => handleAllocationPackageChange(item.packageId, "packageName", event.target.value)} placeholder="Package name" readOnly={allocationEditorMode === "reactivate"} /></td>
                        <td style={{ minWidth: 0 }}><input style={{ width: "100%", maxWidth: "100%", minWidth: 0, boxSizing: "border-box" }} type="number" min="1" step="1" value={item.min} onChange={(event) => handleAllocationPackageChange(item.packageId, "min", event.target.value)} placeholder="1" readOnly={allocationEditorMode === "reactivate"} /></td>
                        <td style={{ minWidth: 0 }}><input style={{ width: "100%", maxWidth: "100%", minWidth: 0, boxSizing: "border-box" }} type="number" min="1" step="1" value={item.max} onChange={(event) => handleAllocationPackageChange(item.packageId, "max", event.target.value)} placeholder={item.noLimit ? "No Limit" : "2"} disabled={item.noLimit || allocationEditorMode === "reactivate"} readOnly={allocationEditorMode === "reactivate"} /></td>
                        <td style={{ textAlign: "center", minWidth: 0 }}><input type="checkbox" checked={item.noLimit} onChange={(event) => handleAllocationPackageChange(item.packageId, "noLimit", event.target.checked)} disabled={allocationEditorMode === "reactivate"} aria-label={`No limit for ${item.packageName || "plan package"}`} /></td>
                        <td style={{ minWidth: 0 }}><input style={{ width: "100%", maxWidth: "100%", minWidth: 0, boxSizing: "border-box" }} type="text" inputMode="decimal" value={item.planFee} onChange={(event) => handleAllocationPackageChange(item.packageId, "planFee", event.target.value)} placeholder="Enter fee" readOnly={allocationEditorMode === "reactivate"} /></td>
                        <td style={{ minWidth: 0 }}><button style={{ width: "100%", maxWidth: "100%", boxSizing: "border-box", whiteSpace: "nowrap" }} type="button" className="erp-settings__secondary-button" onClick={() => handleRemoveAllocationPackage(item.packageId)} disabled={newAllocationModel.packages.length <= 1 || allocationEditorMode === "reactivate"}>Remove</button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="erp-settings__entry-actions" style={{ marginTop: 16, display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button type="button" className="erp-settings__secondary-button" onClick={handleSaveAllocationDraft}>Save Draft</button>
                <button type="button" className="erp-settings__primary-button" onClick={handleActivateAllocationModel}>Activate Model</button>
              </div>
              <p style={{ marginTop: 10 }}>The Effective From date is applied only after <strong>Activate Model</strong> is confirmed.</p>
            </article>
          ) : null}

          {allocationWorkspace === "drafts" ? (
            <article className="erp-settings__card" id="allocation-model-drafts">
              <div className="erp-settings__card-heading">
                <div><h2>Allocation Drafts</h2><p>Saved Allocation drafts remain here until the user explicitly activates them. Drafts may be incomplete and can be edited later.</p></div>
                <button type="button" className="erp-settings__secondary-button" onClick={handleBackToCurrentAllocation}>Back to Current Allocation</button>
              </div>
              {allocationMessage ? <div className="erp-settings__message">{allocationMessage}</div> : null}
              <div className="erp-settings__table-wrapper">
                <table className="erp-settings__table">
                  <thead><tr><th>Chain Name</th><th>Effective From</th><th>Plan Packages</th><th>Status</th><th>Saved</th><th>Action</th></tr></thead>
                  <tbody>
                    {[...allocationDrafts].sort((a, b) => normalizeValue(b.savedAt).localeCompare(normalizeValue(a.savedAt))).map((draft) => (
                      <tr key={draft.draftId}>
                        <td>{draft.name || "Unnamed Draft"}</td>
                        <td>{draft.effectiveFrom || "Not Set"}</td>
                        <td>{Array.isArray(draft.packages) ? draft.packages.map((item) => item.packageName || "Unnamed Package").join(" | ") : "-"}</td>
                        <td>Draft</td>
                        <td>{draft.savedAt ? new Date(draft.savedAt).toLocaleString() : "-"}</td>
                        <td><div className="erp-settings__row-actions"><button type="button" className="erp-settings__table-button" onClick={() => handleEditAllocationDraft(draft)}>Edit</button><button type="button" className="erp-settings__table-button" onClick={() => handleActivateSavedDraft(draft)}>Activate</button><button type="button" className="erp-settings__table-button" onClick={() => handleDeleteAllocationDraft(draft.draftId)}>Delete</button></div></td>
                      </tr>
                    ))}
                    {allocationDrafts.length === 0 ? <tr><td colSpan="6">No Allocation Drafts saved.</td></tr> : null}
                  </tbody>
                </table>
              </div>
            </article>
          ) : null}

          {allocationWorkspace === "history" ? (
            <article className="erp-settings__card" id="allocation-model-history">
              <div className="erp-settings__card-heading">
                <div><h2>Inactive Allocations</h2><p>Allocations moved manually from Active to Inactive remain available here. An Inactive Allocation can be reused or reactivated without overwriting its historical record.</p></div>
                <button type="button" className="erp-settings__secondary-button" onClick={handleBackToCurrentAllocation}>Back to Current Allocation</button>
              </div>
              {allocationMessage ? <div className="erp-settings__message">{allocationMessage}</div> : null}
              <div className="erp-settings__table-wrapper erp-settings__table-wrapper--inactive-allocations">
                <table className="erp-settings__table erp-settings__table--inactive-allocations">
                  <colgroup>
                    <col className="erp-settings__inactive-col-chain" />
                    <col className="erp-settings__inactive-col-from" />
                    <col className="erp-settings__inactive-col-to" />
                    <col className="erp-settings__inactive-col-plan" />
                    <col className="erp-settings__inactive-col-status" />
                    <col className="erp-settings__inactive-col-action" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th>Chain Name</th>
                      <th>Effective From</th>
                      <th>Effective To</th>
                      <th>Plan Fee Structure</th>
                      <th>Status</th>
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {inactiveAllocationModels.map((model) => (
                      <tr key={model.modelId}>
                        <td>{model.name}</td>
                        <td>{model.effectiveFrom || "Legacy"}</td>
                        <td>{model.effectiveTo || "-"}</td>
                        <td>
                          <div className="erp-settings__inactive-plan-structure">
                            {(model.slabs || []).length > 0 ? (
                              model.slabs.map((slab) => (
                                <div key={slab.slabId || `${model.modelId}-${slab.label}`}>
                                  <span>{slab.label}</span>
                                  <span className="erp-settings__inactive-plan-arrow">→</span>
                                  <strong>₹{Number(slab.planFee || 0).toLocaleString("en-IN")}</strong>
                                </div>
                              ))
                            ) : (
                              "-"
                            )}
                          </div>
                        </td>
                        <td className="erp-settings__inactive-status">Inactive</td>
                        <td>
                          <div className="erp-settings__row-actions erp-settings__inactive-actions">
                            <button
                              type="button"
                              className="erp-settings__table-button"
                              onClick={() => handleUseInactiveAllocation(model)}
                            >
                              Use as New
                            </button>
                            <button
                              type="button"
                              className="erp-settings__table-button"
                              onClick={() => handlePrepareInactiveAllocationActivation(model)}
                            >
                              Activate
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                    {inactiveAllocationModels.length === 0 ? <tr><td colSpan="6">No Inactive Allocations available.</td></tr> : null}
                  </tbody>
                </table>
              </div>
            </article>
          ) : null}
        </>      ) : (
      <>

      <article className="erp-settings__card">
        <div className="erp-settings__card-heading">
          <div>
            <h2>Integration Mapping Fields</h2>
            <p>
              Enter one mapping row at a time. New mappings are Active by default.
            </p>
          </div>
        </div>

        <div className="erp-settings__entry-grid">
          <div className="erp-settings__field">
            <label htmlFor="erp-product">Product</label>
            <input
              id="erp-product"
              type="text"
              value={product}
              onChange={(event) => setProduct(event.target.value)}
              placeholder="Billing INC"
            />
          </div>

          <div className="erp-settings__field">
            <label htmlFor="billing-inc-field">Billing INC Field</label>
            <input
              id="billing-inc-field"
              type="text"
              value={billingIncField}
              onChange={(event) => setBillingIncField(event.target.value)}
              placeholder="e.g. GL Account"
            />
          </div>

          <div className="erp-settings__field">
            <label htmlFor="erp-header">ERP Header</label>
            <input
              id="erp-header"
              type="text"
              value={erpHeader}
              onChange={(event) => setErpHeader(event.target.value)}
              placeholder="e.g. G/L Account"
            />
          </div>

          <div className="erp-settings__entry-actions">
            {editingRowKey ? (
              <button
                type="button"
                className="erp-settings__secondary-button"
                onClick={resetEditor}
              >
                Cancel
              </button>
            ) : null}

            <button
              type="button"
              className="erp-settings__primary-button"
              onClick={handleSave}
            >
              {editingRowKey ? "Update ERP Mapping" : "Save ERP Mapping"}
            </button>
          </div>
        </div>

        {saveMessage ? (
          <div className="erp-settings__message">{saveMessage}</div>
        ) : null}
      </article>

      <article className="erp-settings__card">
        <div className="erp-settings__saved-header">
          <div>
            <h2>Saved ERP Mappings</h2>
            <p>
              Active mappings are shown first. Inactive mappings remain below for reference.
            </p>
          </div>

          <div className="erp-settings__search">
            <label htmlFor="erp-mapping-search">Search</label>
            <input
              id="erp-mapping-search"
              type="search"
              value={searchTerm}
              onChange={(event) => setSearchTerm(event.target.value)}
              placeholder="Search product, field, header, or status"
            />
          </div>
        </div>

        {filteredRows.length === 0 ? (
          <div className="erp-settings__empty">
            No ERP mappings match the current search.
          </div>
        ) : (
          <div className="erp-settings__table-wrapper">
            <table className="erp-settings__table">
              <thead>
                <tr>
                  <th>Product</th>
                  <th>Billing INC Field</th>
                  <th>ERP Header</th>
                  <th>Status</th>
                  <th>Action</th>
                </tr>
              </thead>
              <tbody>
                {filteredRows.map((row) => (
                  <tr key={`${row.mappingId}::${row.fieldId}`}>
                    <td>{row.product || "-"}</td>
                    <td>{row.billingIncField || "-"}</td>
                    <td>{row.erpHeader || "-"}</td>
                    <td>{row.status}</td>
                    <td>
                      <div className="erp-settings__row-actions">
                        <button
                          type="button"
                          className="erp-settings__table-button"
                          onClick={() => handleEdit(row)}
                        >
                          Edit
                        </button>
                        <button
                          type="button"
                          className="erp-settings__table-button"
                          onClick={() => handleToggleStatus(row)}
                        >
                          {row.status === "Active" ? "Deactivate" : "Activate"}
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </article>
      </>
      )}
    </section>
  );
}

export default Settings;
