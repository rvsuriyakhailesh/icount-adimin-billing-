export const GOVERNANCE_EVENT_CATEGORIES = [
  "Allocation / Reallocation",
  "Billing",
  "Commercial Terms",
  "Incentive",
  "Invoice",
  "Payments",
  "Record Correction",
  "Site Status",
  "Stage Movement",
];

export function getGovernanceEventCategory(label) {
  const value = String(label ?? "").trim().toLowerCase();
  if (!value) return null;

  if (/allocation|reallocation/.test(value)) {
    return "Allocation / Reallocation";
  }
  if (/record correction|record corrected|correction completed|correction/.test(value)) {
    return "Record Correction";
  }
  if (/invoice|replacement invoice/.test(value)) {
    return "Invoice";
  }
  if (/incentive/.test(value)) {
    return "Incentive";
  }
  if (/price change|subscription fee|subscription mode|subscription type|pricing method|commercial|otf/.test(value)) {
    return "Commercial Terms";
  }
  if (/payment|receipt/.test(value)) {
    return "Payments";
  }
  if (/site inactive|site inactivated|site paused|site activated|service status/.test(value)) {
    return "Site Status";
  }
  if (/entered stage|moved to stage|returned to stage|moved to billings|stage movement/.test(value)) {
    return "Stage Movement";
  }
  if (/billing/.test(value)) {
    return "Billing";
  }

  return null;
}
