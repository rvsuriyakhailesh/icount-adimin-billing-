function normalizeWorkflowValue(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[-_]/g, " ")
    .replace(/\s+/g, " ");
}

// The top-level Site workflow fields are authoritative. Stage snapshots are
// historical records and must never determine the current workflow stage.
export function getCurrentWorkflowStage(record) {
  const processingStatus = normalizeWorkflowValue(record?.processingStatus);

  if (processingStatus === "stage 1") return "Stage 1";
  if (processingStatus === "stage 2") return "Stage 2";
  if (processingStatus === "stage 3") return "Stage 3";

  const readinessStatus = normalizeWorkflowValue(record?.readinessStatus);

  if (["ready", "returned to stage 1"].includes(readinessStatus)) {
    return "Stage 1";
  }
  if (readinessStatus === "moved to stage 2") return "Stage 2";
  if (readinessStatus === "moved to stage 3") return "Stage 3";

  return "";
}

export function isStage1ReadyForSelection(record) {
  const readinessStatus = normalizeWorkflowValue(record?.readinessStatus);

  return (
    getCurrentWorkflowStage(record) === "Stage 1" &&
    ["ready", "returned to stage 1"].includes(readinessStatus)
  );
}

export function isCurrentDownstreamWorkflowStage(record) {
  return ["Stage 2", "Stage 3"].includes(getCurrentWorkflowStage(record));
}
