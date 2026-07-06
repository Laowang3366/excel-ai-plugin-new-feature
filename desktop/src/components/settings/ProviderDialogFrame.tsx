import React from "react";
import { X } from "../common/IconMap";

interface ProviderDialogFrameProps {
  dialogClassName: string;
  title: string;
  onClose: () => void;
  children: React.ReactNode;
  actions: React.ReactNode;
}

export function ProviderDialogFrame({
  dialogClassName,
  title,
  onClose,
  children,
  actions,
}: ProviderDialogFrameProps) {
  return (
    <div className="dialog-overlay" onClick={onClose}>
      <div className={`dialog ${dialogClassName}`} onClick={(event) => event.stopPropagation()}>
        <div className="dialog-header">
          <h3>{title}</h3>
          <button className="dialog-close" onClick={onClose}><X size={16} /></button>
        </div>

        <div className="dialog-body">
          {children}
        </div>

        <div className="dialog-actions">
          {actions}
        </div>
      </div>
    </div>
  );
}
