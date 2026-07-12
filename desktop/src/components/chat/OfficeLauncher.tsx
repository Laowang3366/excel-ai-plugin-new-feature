import { useEffect, useRef, useState } from "react";

import type { OfficeApplication } from "../../electronApi";
import { ipcApi } from "../../services/ipcApi";
import {
  ChevronDown,
  FileBarChart,
  FileSpreadsheet,
  FileText,
} from "../common/IconMap";

interface OfficeLauncherText {
  open: (application: string) => string;
  choose: string;
  failed: string;
}

interface OfficeLauncherProps {
  text: OfficeLauncherText;
}

const APPLICATION_LABELS: Record<OfficeApplication, string> = {
  wps: "WPS",
  excel: "Microsoft Excel",
  word: "Microsoft Word",
  powerpoint: "Microsoft PowerPoint",
};

const APPLICATIONS = Object.keys(APPLICATION_LABELS) as OfficeApplication[];

export function OfficeLauncher({ text }: OfficeLauncherProps) {
  const [selectedApplication, setSelectedApplication] = useState<OfficeApplication>("wps");
  const [menuOpen, setMenuOpen] = useState(false);
  const [launching, setLaunching] = useState(false);
  const [error, setError] = useState("");
  const rootRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!menuOpen) return;

    const closeOnOutsideClick = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setMenuOpen(false);
    };
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === "Escape") setMenuOpen(false);
    };

    document.addEventListener("pointerdown", closeOnOutsideClick);
    document.addEventListener("keydown", closeOnEscape);
    return () => {
      document.removeEventListener("pointerdown", closeOnOutsideClick);
      document.removeEventListener("keydown", closeOnEscape);
    };
  }, [menuOpen]);

  const launch = async (application: OfficeApplication) => {
    setSelectedApplication(application);
    setLaunching(true);
    setError("");
    try {
      const result = await ipcApi.app.launchOffice(application);
      if (result.success) {
        setMenuOpen(false);
      } else {
        setError(result.error || text.failed);
        setMenuOpen(true);
      }
    } catch (launchError) {
      setError(launchError instanceof Error ? launchError.message : text.failed);
      setMenuOpen(true);
    } finally {
      setLaunching(false);
    }
  };

  return (
    <div className="office-launcher" ref={rootRef}>
      <div className="office-launcher-split">
        <button
          className="office-launcher-primary"
          type="button"
          disabled={launching}
          title={text.open(APPLICATION_LABELS[selectedApplication])}
          onClick={() => void launch(selectedApplication)}
        >
          <span className={`office-launcher-app-icon ${selectedApplication}`}>
            {getApplicationIcon(selectedApplication)}
          </span>
          <span className="office-launcher-app-label">
            {APPLICATION_LABELS[selectedApplication]}
          </span>
        </button>
        <button
          className={`office-launcher-toggle${menuOpen ? " active" : ""}`}
          type="button"
          title={text.choose}
          aria-label={text.choose}
          aria-haspopup="menu"
          aria-expanded={menuOpen}
          onClick={() => {
            setError("");
            setMenuOpen((open) => !open);
          }}
        >
          <ChevronDown size={14} />
        </button>
      </div>

      {menuOpen && (
        <div className="office-launcher-menu" role="menu" aria-label={text.choose}>
          {APPLICATIONS.map((application) => (
            <button
              key={application}
              className={`office-launcher-item ${application}`}
              type="button"
              role="menuitem"
              disabled={launching}
              onClick={() => void launch(application)}
            >
              {getApplicationIcon(application)}
              <span>{APPLICATION_LABELS[application]}</span>
            </button>
          ))}
          {error && <div className="office-launcher-error">{error}</div>}
        </div>
      )}
    </div>
  );
}

function getApplicationIcon(application: OfficeApplication) {
  if (application === "word") return <FileText size={17} />;
  if (application === "powerpoint") return <FileBarChart size={17} />;
  return <FileSpreadsheet size={17} />;
}
