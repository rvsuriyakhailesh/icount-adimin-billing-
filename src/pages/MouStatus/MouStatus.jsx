import { useMemo, useState } from "react";
import "./MouStatus.css";

const FILTERS = ["All", "Live", "About to Expire", "Expired", "Renewal Pending"];

function normalize(value) {
  return String(value ?? "").trim();
}

function parseDate(value) {
  const text = normalize(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
  const [year, month, day] = text.split("-").map(Number);
  const date = new Date(year, month - 1, day);
  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) return null;
  return date;
}

function getDaysRemaining(mouEndDate) {
  const end = parseDate(mouEndDate);
  if (!end) return null;

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  return Math.max(
    0,
    Math.ceil((end.getTime() - today.getTime()) / (1000 * 60 * 60 * 24)),
  );
}

function getDisplayMouStatus(record) {
  const explicit = normalize(record.mouStatus);
  const days = getDaysRemaining(record.mouEndDate);

  if (days === 0 && parseDate(record.mouEndDate)) return "Expired";
  if (explicit) return explicit;

  return record.mouReceivedDate ? "Live" : "Sent";
}

function getMonitoringStatus(record) {
  const mouStatus = getDisplayMouStatus(record);
  const days = getDaysRemaining(record.mouEndDate);
  const renewal = normalize(record.renewalStatus).toLowerCase();

  if (renewal === "renewal pending") return "Renewal Pending";
  if (mouStatus.toLowerCase() === "expired" || days === 0) return "Expired";
  if (days !== null && days <= 30) return "About to Expire";

  return "Live";
}

function MouStatus({
  records = [],
  onOpen = () => {},
  onCreateAddendum = () => {},
  onOpenRenewal = () => {},
}) {
  const [activeFilter, setActiveFilter] = useState("All");
  const [search, setSearch] = useState("");
  const [selectedRecord, setSelectedRecord] = useState(null);

  const rows = useMemo(
    () =>
      records
        .filter((record) => normalize(record.screenCode || record.siteId))
        .map((record, index) => ({
          ...record,
          _key:
            normalize(record.recordId) ||
            normalize(record.stage1GroupId) ||
            normalize(record.screenCode) ||
            `mou-${index}`,
          _customerCode:
            normalize(record.billingCode) ||
            normalize(record.customerCode) ||
            normalize(record.billingId),
          _screenCode: normalize(record.screenCode) || normalize(record.siteId),
          _screenName:
            normalize(record.screenName) || normalize(record.erpScreenName),
          _location: normalize(record.location) || "-",
          _mouStatus: getDisplayMouStatus(record),
          _daysRemaining: getDaysRemaining(record.mouEndDate),
          _monitoringStatus: getMonitoringStatus(record),
        })),
    [records],
  );

  const counts = useMemo(() => {
    const next = Object.fromEntries(FILTERS.map((filter) => [filter, 0]));
    next.All = rows.length;

    rows.forEach((row) => {
      if (next[row._monitoringStatus] !== undefined) {
        next[row._monitoringStatus] += 1;
      }
    });

    return next;
  }, [rows]);

  const visibleRows = useMemo(() => {
    const query = search.trim().toLowerCase();

    return rows.filter((row) => {
      if (activeFilter !== "All" && row._monitoringStatus !== activeFilter) {
        return false;
      }

      if (!query) return true;

      return [
        row._customerCode,
        row._screenCode,
        row._screenName,
        row._location,
        row._mouStatus,
      ]
        .join(" ")
        .toLowerCase()
        .includes(query);
    });
  }, [activeFilter, rows, search]);

  function handleOpen(record) {
    setSelectedRecord(record);
    onOpen(record);
  }

  function handlePrimaryAction(record) {
    if (record._monitoringStatus === "Live") {
      onCreateAddendum(record);
      return;
    }

    onOpenRenewal(record);
  }

  return (
    <section className="mou-status-page">
      <div className="mou-status-page__header">
        <span className="mou-status-page__eyebrow">MOU STATUS</span>
        <h2>MoU Status</h2>
        <p>Internal overview of MoU expiry and renewal status across all sites.</p>
      </div>

      <div className="mou-status-page__toolbar-card">
        <div className="mou-status-page__tabs">
          {FILTERS.map((filter) => (
            <button
              key={filter}
              type="button"
              className={`mou-status-page__tab ${
                activeFilter === filter ? "mou-status-page__tab--active" : ""
              }`}
              onClick={() => setActiveFilter(filter)}
            >
              <span>{filter}</span>
              <strong>{counts[filter] || 0}</strong>
            </button>
          ))}
        </div>

        <label className="mou-status-page__search">
          <span>Search</span>
          <input
            type="search"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            placeholder="Customer Code, Screen Code, Screen Name, Location"
          />
        </label>
      </div>

      <div className="mou-status-page__table-card">
        <div className="mou-status-page__table-wrap">
          <table className="mou-status-page__table">
            <colgroup>
              <col style={{ width: "15%" }} />
              <col style={{ width: "14%" }} />
              <col style={{ width: "22%" }} />
              <col style={{ width: "19%" }} />
              <col style={{ width: "12%" }} />
              <col style={{ width: "10%" }} />
              <col style={{ width: "8%" }} />
            </colgroup>
            <thead>
              <tr>
                <th>Customer Code</th>
                <th>Screen Code</th>
                <th>Screen Name</th>
                <th>Location</th>
                <th>MoU Status</th>
                <th>Days Remaining</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {visibleRows.length === 0 ? (
                <tr>
                  <td colSpan="7" className="mou-status-page__empty">
                    No MoU records found for the selected view.
                  </td>
                </tr>
              ) : (
                visibleRows.map((record) => (
                  <tr key={record._key}>
                    <td>{record._customerCode || "-"}</td>
                    <td>{record._screenCode || "-"}</td>
                    <td className="mou-status-page__ellipsis">
                      {record._screenName || "-"}
                    </td>
                    <td className="mou-status-page__ellipsis">
                      {record._location}
                    </td>
                    <td>
                      <span
                        className={`mou-status-page__status mou-status-page__status--${record._monitoringStatus
                          .toLowerCase()
                          .replace(/\s+/g, "-")}`}
                      >
                        {record._monitoringStatus === "About to Expire"
                          ? "About to Expire"
                          : record._mouStatus}
                      </span>
                    </td>
                    <td>
                      <span
                        className={`mou-status-page__days ${
                          record._daysRemaining !== null &&
                          record._daysRemaining <= 30
                            ? "mou-status-page__days--attention"
                            : ""
                        }`}
                      >
                        {record._daysRemaining ?? "—"}
                      </span>
                    </td>
                    <td>
                      <button
                        type="button"
                        className="mou-status-page__open"
                        onClick={() => handleOpen(record)}
                      >
                        Open
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        <div className="mou-status-page__footer">
          Showing {visibleRows.length} of {rows.length} records
        </div>
      </div>

      {selectedRecord && (
        <div className="mou-status-page__details-card">
          <div className="mou-status-page__details-header">
            <div>
              <span className="mou-status-page__details-kicker">SELECTED MOU</span>
              <h3>
                {selectedRecord._screenCode} / {selectedRecord._screenName}
              </h3>
            </div>

            <button
              type="button"
              className="mou-status-page__details-close"
              onClick={() => setSelectedRecord(null)}
              aria-label="Close MoU details"
            >
              ×
            </button>
          </div>

          <div className="mou-status-page__details-grid">
            <div>
              <span>Complex Code</span>
              <strong>{normalize(selectedRecord.complexCode) || "-"}</strong>
            </div>
            <div>
              <span>Location</span>
              <strong>{selectedRecord._location}</strong>
            </div>
            <div>
              <span>MoU Start Date</span>
              <strong>{normalize(selectedRecord.mouStartDate) || "-"}</strong>
            </div>
            <div>
              <span>MoU End Date</span>
              <strong>{normalize(selectedRecord.mouEndDate) || "-"}</strong>
            </div>
            <div>
              <span>Agreement Type</span>
              <strong>{normalize(selectedRecord.agreementType) || "MoU"}</strong>
            </div>
            <div>
              <span>Renewal Status</span>
              <strong>{selectedRecord._monitoringStatus}</strong>
            </div>
          </div>

          <div className="mou-status-page__details-actions">
            <button
              type="button"
              className="mou-status-page__details-secondary"
              onClick={() => setSelectedRecord(null)}
            >
              Close
            </button>

            <button
              type="button"
              className="mou-status-page__details-primary"
              onClick={() => handlePrimaryAction(selectedRecord)}
            >
              {selectedRecord._monitoringStatus === "Live"
                ? "Create Addendum"
                : selectedRecord._monitoringStatus === "Renewal Pending"
                  ? "Continue Renewal"
                  : "Open Renewal"}
            </button>
          </div>
        </div>
      )}
    </section>
  );
}

export default MouStatus;
