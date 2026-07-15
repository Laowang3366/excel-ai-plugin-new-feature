import type { OfficeAutomationTransaction } from "../../electronApi";
import { AlertTriangle, Redo2, Undo2 } from "../common/IconMap";
import { formatOfficeTime, officeStatusLabel, shortOfficePath } from "./officeAutomationViewModel";
import { EmptyState, StatusBadge, Toolbar } from "./OfficeAutomationPanelShared";

type ForceRequest = { action: "undo" | "redo"; transaction: OfficeAutomationTransaction };

export type OfficeAutomationTransactionsTabProps = {
  state: {
    transactions: OfficeAutomationTransaction[];
    selectedTransactionId: string;
    forceRequest: ForceRequest | null;
    busy: string;
  };
  actions: {
    setSelectedTransactionId: (id: string) => void;
    setForceRequest: (value: ForceRequest | null) => void;
    refresh: () => void;
    applyTransaction: (
      action: "undo" | "redo",
      transaction: OfficeAutomationTransaction,
      force?: boolean,
    ) => void;
  };
};

function canUndo(transaction: OfficeAutomationTransaction): boolean {
  return (
    transaction.status === "applied" ||
    (transaction.status === "conflicted" && transaction.conflictBaseStatus !== "undone")
  );
}

function canRedo(transaction: OfficeAutomationTransaction): boolean {
  return (
    transaction.status === "undone" ||
    (transaction.status === "conflicted" && transaction.conflictBaseStatus === "undone")
  );
}

export function OfficeAutomationTransactionsTab({
  state,
  actions,
}: OfficeAutomationTransactionsTabProps) {
  const selectedTransaction = state.transactions.find(
    (item) => item.id === state.selectedTransactionId,
  );

  return (
    <div className="office-automation-view">
      <Toolbar
        title="事务与恢复"
        count={state.transactions.length}
        onRefresh={actions.refresh}
        busy={state.busy === "refresh:transactions"}
      />
      <div className="office-master-detail">
        <div className="office-master-list">
          {state.transactions.map((transaction) => (
            <button
              key={transaction.id}
              type="button"
              className={transaction.id === state.selectedTransactionId ? "selected" : ""}
              onClick={() => actions.setSelectedTransactionId(transaction.id)}
            >
              <span className={`office-status-dot ${transaction.status}`} />
              <span>
                <strong>{transaction.changes[0]?.detail || "Office 事务"}</strong>
                <small>
                  {officeStatusLabel(transaction.status)} · {transaction.changes.length} 项修改
                </small>
              </span>
            </button>
          ))}
          {state.transactions.length === 0 && <EmptyState text="暂无事务记录" />}
        </div>
        <div className="office-detail-pane">
          {selectedTransaction ? (
            <>
              <div className="office-detail-title">
                <div>
                  <strong>修改清单</strong>
                  <small>
                    {formatOfficeTime(selectedTransaction.updatedAt)} ·{" "}
                    {selectedTransaction.id.slice(0, 8)}
                  </small>
                </div>
                <StatusBadge status={selectedTransaction.status} />
              </div>
              <div className="office-detail-actions">
                {canUndo(selectedTransaction) && (
                  <button
                    type="button"
                    className="office-command"
                    disabled={state.busy === `undo:${selectedTransaction.id}`}
                    onClick={() => void actions.applyTransaction("undo", selectedTransaction)}
                  >
                    <Undo2 size={14} />
                    撤销
                  </button>
                )}
                {canRedo(selectedTransaction) && (
                  <button
                    type="button"
                    className="office-command"
                    disabled={state.busy === `redo:${selectedTransaction.id}`}
                    onClick={() => void actions.applyTransaction("redo", selectedTransaction)}
                  >
                    <Redo2 size={14} />
                    重做
                  </button>
                )}
              </div>
              {selectedTransaction.error && (
                <div className="office-inline-error">{selectedTransaction.error}</div>
              )}
              {state.forceRequest?.transaction.id === selectedTransaction.id && (
                <div className="office-force-confirm">
                  <AlertTriangle size={15} />
                  <span>
                    {state.forceRequest.transaction.conflicts
                      ?.map((item) => `${shortOfficePath(item.filePath)}：${item.reason}`)
                      .join("；") || "文件已在事务外修改"}
                  </span>
                  <div>
                    <button type="button" onClick={() => actions.setForceRequest(null)}>
                      取消
                    </button>
                    <button
                      type="button"
                      className="danger"
                      onClick={() =>
                        void actions.applyTransaction(
                          state.forceRequest!.action,
                          state.forceRequest!.transaction,
                          true,
                        )
                      }
                    >
                      确认覆盖
                    </button>
                  </div>
                </div>
              )}
              <div className="office-change-list">
                {selectedTransaction.changes.map((change, index) => (
                  <div key={`${change.kind}:${index}`}>
                    <strong>{change.detail}</strong>
                    <small>{change.target || change.kind}</small>
                  </div>
                ))}
                {selectedTransaction.changes.length === 0 && (
                  <EmptyState text="没有记录到文件修改" />
                )}
              </div>
            </>
          ) : (
            <EmptyState text="选择事务查看修改清单" />
          )}
        </div>
      </div>
    </div>
  );
}
