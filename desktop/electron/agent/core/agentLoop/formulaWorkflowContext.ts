import type { Turn, TurnItem } from "../../shared/types";
import { getLatestFormulaPreparation, isFormulaWorkflowTurn } from "./formulaTaskContract";
import { getLatestFormulaVerification } from "./formulaVerification";

export function createFormulaWorkflowTurnView(
  currentTurn: Turn,
  itemGroups: TurnItem[][],
): Turn {
  if (isFormulaWorkflowTurn(currentTurn)) return currentTurn;
  const historyItems = findPendingFormulaWorkflowItems(itemGroups, currentTurn.items);
  if (!historyItems) return currentTurn;

  const view = { ...currentTurn } as Turn;
  Object.defineProperty(view, "items", {
    enumerable: true,
    get: () => [...historyItems, ...currentTurn.items],
  });
  return view;
}

export function hasPendingFormulaWorkflowHistory(
  itemGroups: TurnItem[][],
  currentItems?: TurnItem[],
): boolean {
  return Boolean(findPendingFormulaWorkflowItems(itemGroups, currentItems));
}

function findPendingFormulaWorkflowItems(
  itemGroups: TurnItem[][],
  currentItems?: TurnItem[],
): TurnItem[] | null {
  for (let index = itemGroups.length - 1; index >= 0; index--) {
    const items = itemGroups[index];
    if (items === currentItems || sameCurrentGroup(items, currentItems)) continue;
    const candidate: Turn = {
      threadId: "formula-history",
      turnId: `formula-history-${index}`,
      status: "completed",
      startedAt: 0,
      items,
    };
    const preparation = getLatestFormulaPreparation(candidate)?.preparation;
    const verification = getLatestFormulaVerification(candidate);
    const hasWorkflowActivity = isFormulaWorkflowTurn(candidate)
      || Boolean(preparation)
      || Boolean(verification)
      || items.some((item) => item.type === "tool_call" && item.toolName === "range.write");
    if (!hasWorkflowActivity) continue;
    if (verification?.status === "passed" || verification?.status === "passed_with_assumptions") return null;
    if (!preparation && verification?.status !== "failed") return null;

    let rootIndex = index;
    for (let previous = index; previous >= 0; previous--) {
      const previousTurn: Turn = {
        threadId: "formula-history",
        turnId: `formula-history-${previous}`,
        status: "completed",
        startedAt: 0,
        items: itemGroups[previous],
      };
      if (isFormulaWorkflowTurn(previousTurn)) {
        rootIndex = previous;
        break;
      }
    }
    return itemGroups.slice(rootIndex, index + 1).flat();
  }
  return null;
}

function sameCurrentGroup(items: TurnItem[], currentItems?: TurnItem[]): boolean {
  if (!currentItems || items.length !== currentItems.length || items.length === 0) return false;
  return items[0] === currentItems[0] && items[items.length - 1] === currentItems[currentItems.length - 1];
}
