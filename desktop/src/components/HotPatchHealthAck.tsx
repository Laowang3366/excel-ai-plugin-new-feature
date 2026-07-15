import { useEffect } from "react";

import { ipcApi } from "../services/ipcApi";

export function HotPatchHealthAck() {
  useEffect(() => {
    let canceled = false;
    const frame = window.requestAnimationFrame(() => {
      window.requestAnimationFrame(() => {
        if (canceled || document.querySelector(".error-boundary")) return;
        void ipcApi.update.ackHotPatchHealth();
      });
    });
    return () => {
      canceled = true;
      window.cancelAnimationFrame(frame);
    };
  }, []);
  return null;
}
