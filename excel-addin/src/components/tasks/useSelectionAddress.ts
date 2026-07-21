import { useCallback, useState } from "react";
import type { HostAdapter } from "@shared/host";

export function formatSelectionAddress(sheetName: string, address: string): string {
  const sheet = sheetName.trim();
  const addr = address.trim();
  if (!sheet) return addr;
  if (!addr) return sheet;
  if (addr.includes("!")) return addr;
  const needsQuote = /[\s'!]/.test(sheet);
  const quoted = needsQuote ? `'${sheet.replace(/'/g, "''")}'` : sheet;
  return `${quoted}!${addr}`;
}

export function useSelectionAddress(adapter: HostAdapter | null) {
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | undefined>();

  const readSelection = useCallback(async (): Promise<string | null> => {
    if (!adapter) {
      setError("宿主未就绪");
      return null;
    }
    setBusy(true);
    setError(undefined);
    try {
      const result = await adapter.getSelection();
      if (!result.ok) {
        setError(result.reason || "读取选区失败");
        return null;
      }
      const formatted = formatSelectionAddress(
        result.data.sheetName,
        result.data.address,
      );
      return formatted || null;
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      return null;
    } finally {
      setBusy(false);
    }
  }, [adapter]);

  return { readSelection, busy, error, setError };
}
