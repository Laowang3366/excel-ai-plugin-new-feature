import { useState } from "react";
import type { HostAdapter, HostStatus } from "@shared/host";

interface Props {
  adapter: HostAdapter;
}

export function HostStatusPanel({ adapter }: Props) {
  const [status, setStatus] = useState<HostStatus | null>(null);
  const [error, setError] = useState<string | null>(null);

  async function refresh() {
    setError(null);
    const result = await adapter.getStatus();
    if (!result.ok) {
      setStatus(null);
      setError(result.reason);
      return;
    }
    setStatus(result.data);
  }

  return (
    <section className="card">
      <h2>宿主状态</h2>
      <div className="row">
        <button type="button" onClick={() => void refresh()}>
          刷新连接
        </button>
      </div>
      {error && <p className="muted">错误：{error}</p>}
      {status && (
        <pre>{JSON.stringify(status, null, 2)}</pre>
      )}
      {!status && !error && (
        <p className="muted">点击刷新以读取 workbook / Application 状态。</p>
      )}
    </section>
  );
}
