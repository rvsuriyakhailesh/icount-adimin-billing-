import { cloneElement, isValidElement, useEffect, useMemo, useState } from "react";
import "./AppLayout.css";
import Sidebar from "./Sidebar";

const pageTitles = {
  "site-entry": "Stage 1 Site Entry",
  "mou-status": "MoU Status",
  installations: "Stage 2 Installations",
  billings: "Stage 3 Billings",
  incentive: "Incentive",
  history: "History Tracking",
  audit: "Audit History",
  downloads: "Downloads",
  settings: "Settings",
};

function AppLayout({ activePage, onPageChange, children }) {
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [minimizedPages, setMinimizedPages] = useState({
    "site-entry": false,
    installations: false,
  });

  const isStagePage = useMemo(
    () => activePage === "site-entry" || activePage === "installations",
    [activePage],
  );

  const isMinimized = isStagePage ? minimizedPages[activePage] || false : false;

  useEffect(() => {
    function handleFullscreenChange() {
      const active = Boolean(document.fullscreenElement);

      console.log("[Fullscreen change]", {
        active,
        element: document.fullscreenElement,
        activePage,
        timestamp: new Date().toISOString(),
      });

      setIsFullscreen(active && document.fullscreenElement === document.documentElement);
    }

    document.addEventListener("fullscreenchange", handleFullscreenChange);
    setIsFullscreen(document.fullscreenElement === document.documentElement);

    return () => {
      document.removeEventListener("fullscreenchange", handleFullscreenChange);
    };
  }, [activePage]);

  function requestBrowserFullscreen() {
    const requestPromise = document.documentElement.requestFullscreen?.();
    requestPromise?.catch?.(() => {});
    return requestPromise;
  }

  function exitBrowserFullscreen() {
    const exitPromise = document.exitFullscreen?.();
    exitPromise?.catch?.(() => {});
    return exitPromise;
  }

  function handleToggleFullscreen() {
    if (document.fullscreenElement === document.documentElement) {
      exitBrowserFullscreen();
      return;
    }

    requestBrowserFullscreen();
  }

  function handleToggleMinimized() {
    if (!isStagePage) {
      return;
    }

    setMinimizedPages((currentPages) => ({
      ...currentPages,
      [activePage]: !isMinimized,
    }));
  }

  function handleMinimizedChange(nextMinimized) {
    if (!isStagePage) {
      return;
    }

    setMinimizedPages((currentPages) => ({
      ...currentPages,
      [activePage]: Boolean(nextMinimized),
    }));
  }

  const renderedChildren =
    isStagePage && isValidElement(children)
      ? cloneElement(children, {
          isMinimized,
          isFullscreenActive: isFullscreen,
          onRequestFullscreenRestore: requestBrowserFullscreen,
          onMinimizedChange: handleMinimizedChange,
        })
      : children;

  return (
    <div className="app-layout">
      <Sidebar activePage={activePage} onPageChange={onPageChange} />

      <main className="app-layout__main">
        <header className="app-layout__header">
          <div>
            <p className="app-layout__eyebrow">Billing and INC</p>
            <h2>{pageTitles[activePage]}</h2>
          </div>

          {isStagePage ? (
            <div className="app-layout__header-actions">
              <button
                type="button"
                className="app-layout__header-button"
                onClick={handleToggleMinimized}
              >
                {isMinimized ? "Expand" : "Minimize"}
              </button>

              <button
                type="button"
                className="app-layout__header-button app-layout__header-button--primary"
                onClick={handleToggleFullscreen}
              >
                {isFullscreen ? "Exit Full Screen" : "Full Screen"}
              </button>
            </div>
          ) : null}
        </header>

        <section className="app-layout__content">{renderedChildren}</section>
      </main>
    </div>
  );
}

export default AppLayout;
