import "./Sidebar.css";

const navigationItems = [
  { id: "site-entry", label: "Site Entry", icon: "▣" },
  { id: "installations", label: "Installations", icon: "⌁" },
  { id: "billings", label: "Billings", icon: "₹" },
  { id: "incentive", label: "Incentive", icon: "★" },
  { id: "history", label: "History Tracking", icon: "↺" },
  { id: "audit", label: "Audit History", icon: "✓" },
  { id: "downloads", label: "Downloads", icon: "⇩" },
  { id: "settings", label: "Settings", icon: "⚙" },
];

function Sidebar({ activePage, onPageChange }) {
  return (
    <aside className="sidebar">
      <div className="sidebar__brand">
        <div className="sidebar__brand-mark">B</div>

        <div>
          <h1>Billing and INC</h1>
          <p>Operations Portal</p>
        </div>
      </div>

      <nav className="sidebar__navigation" aria-label="Main navigation">
        <p className="sidebar__section-label">Workflow</p>

        {navigationItems.slice(0, 4).map((item) => (
          <button
            key={item.id}
            type="button"
            className={`sidebar__item ${
              activePage === item.id ? "sidebar__item--active" : ""
            }`}
            onClick={() => onPageChange(item.id)}
          >
            <span className="sidebar__icon" aria-hidden="true">
              {item.icon}
            </span>

            <span>{item.label}</span>
          </button>
        ))}

        <p className="sidebar__section-label sidebar__section-label--spaced">
          Governance
        </p>

        {navigationItems.slice(4).map((item) => (
          <button
            key={item.id}
            type="button"
            className={`sidebar__item ${
              activePage === item.id ? "sidebar__item--active" : ""
            }`}
            onClick={() => onPageChange(item.id)}
          >
            <span className="sidebar__icon" aria-hidden="true">
              {item.icon}
            </span>

            <span>{item.label}</span>
          </button>
        ))}
      </nav>

      <div className="sidebar__footer">
        <span className="sidebar__status-dot" />
        Frontend Development
      </div>
    </aside>
  );
}

export default Sidebar;
