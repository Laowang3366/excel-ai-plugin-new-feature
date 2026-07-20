/**
 * Text-parse dependency graph (desktop DependencyReport semantics without COM).
 */

import { makeCellId, normalizeA1Address } from "./address";
import { parseFormulaReferences, referencesToEdges } from "./references";
import type {
  BrokenReference,
  FormulaCellRecord,
  FormulaDependencyNode,
  FormulaDependencyReport,
  FormulaEdge,
} from "./types";
import { DEPENDENCY_LIMITATIONS } from "./types";

function isFormulaText(text: string): boolean {
  return text.length > 1 && text.startsWith("=");
}

function dedupeEdges(edges: FormulaEdge[]): FormulaEdge[] {
  const seen = new Set<string>();
  const out: FormulaEdge[] = [];
  for (const edge of edges) {
    const key = `${edge.from.toLowerCase()}|${edge.to.toLowerCase()}|${edge.kind}|${edge.reference}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(edge);
  }
  return out;
}

/**
 * DFS cycle detection restricted to edges whose `to` is also a formula node
 * (desktop FindCycles).
 */
export function findCycles(
  nodeIds: Iterable<string>,
  edges: FormulaEdge[],
): string[][] {
  const idSet = new Map<string, string>();
  for (const id of nodeIds) idSet.set(id.toLowerCase(), id);

  const graph = new Map<string, string[]>();
  for (const edge of edges) {
    const fromKey = edge.from.toLowerCase();
    const toKey = edge.to.toLowerCase();
    if (!idSet.has(toKey)) continue;
    const from = idSet.get(fromKey);
    const to = idSet.get(toKey);
    if (!from || !to) continue;
    const list = graph.get(from) ?? [];
    list.push(to);
    graph.set(from, list);
  }

  const states = new Map<string, 0 | 1 | 2>();
  const stack: string[] = [];
  const cycles: string[][] = [];

  const visit = (node: string): void => {
    states.set(node.toLowerCase(), 1);
    stack.push(node);
    for (const next of graph.get(node) ?? []) {
      const st = states.get(next.toLowerCase());
      if (st === undefined) {
        visit(next);
      } else if (st === 1) {
        const start = stack.findIndex(
          (item) => item.toLowerCase() === next.toLowerCase(),
        );
        if (start >= 0) {
          cycles.push([...stack.slice(start), next]);
        }
      }
    }
    stack.pop();
    states.set(node.toLowerCase(), 2);
  };

  for (const id of idSet.values()) {
    if (!states.has(id.toLowerCase())) visit(id);
  }

  const seenPath = new Set<string>();
  const unique: string[][] = [];
  for (const path of cycles) {
    const key = path.map((p) => p.toLowerCase()).join("->");
    if (seenPath.has(key)) continue;
    seenPath.add(key);
    unique.push(path);
  }
  return unique;
}

export function buildDependencyReport(
  cells: FormulaCellRecord[],
): FormulaDependencyReport {
  const formulaCells = cells.filter((c) => isFormulaText(c.formula));
  const edges: FormulaEdge[] = [];
  const broken: BrokenReference[] = [];
  const rawNodes: {
    id: string;
    sheet: string;
    address: string;
    formula: string;
    value?: unknown;
  }[] = [];

  for (const cell of formulaCells) {
    const address = normalizeA1Address(cell.address);
    const id = makeCellId(cell.sheetName, address);
    rawNodes.push({
      id,
      sheet: cell.sheetName.trim(),
      address,
      formula: cell.formula,
      value: cell.value,
    });
    if (cell.formula.toUpperCase().includes("#REF!")) {
      broken.push({ cell: id, formula: cell.formula, reason: "#REF!" });
    }
    const refs = parseFormulaReferences(cell.formula, cell.sheetName.trim());
    edges.push(...referencesToEdges(id, refs));
  }

  const uniqueEdges = dedupeEdges(edges);
  const nodeIds = rawNodes.map((n) => n.id);
  const nodeIdSet = new Set(nodeIds.map((id) => id.toLowerCase()));

  const precedents = new Map<string, string[]>();
  const dependents = new Map<string, string[]>();

  const pushMap = (map: Map<string, string[]>, key: string, value: string) => {
    const k = key.toLowerCase();
    const list = map.get(k) ?? [];
    if (!list.some((v) => v.toLowerCase() === value.toLowerCase())) {
      list.push(value);
    }
    map.set(k, list);
  };

  for (const edge of uniqueEdges) {
    pushMap(precedents, edge.from, edge.to);
    if (nodeIdSet.has(edge.to.toLowerCase())) {
      pushMap(dependents, edge.to, edge.from);
    }
  }

  const nodes: FormulaDependencyNode[] = rawNodes.map((n) => ({
    id: n.id,
    sheet: n.sheet,
    address: n.address,
    formula: n.formula,
    value: n.value,
    precedents: precedents.get(n.id.toLowerCase()) ?? [],
    dependents: dependents.get(n.id.toLowerCase()) ?? [],
  }));

  const cycles = findCycles(nodeIds, uniqueEdges).map((path) => ({ path }));

  return {
    nodes,
    edges: uniqueEdges,
    cycles,
    brokenReferences: broken,
    formulaCount: nodes.length,
    edgeCount: uniqueEdges.length,
    limitations: DEPENDENCY_LIMITATIONS,
  };
}
