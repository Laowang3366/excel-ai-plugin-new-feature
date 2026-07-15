import * as fs from "fs";
import * as path from "path";

import type { ThreadId } from "../shared/types";
import { getDefaultSessionsRoot } from "./sessionStore";

export type ThreadSpawnEdgeStatus = "open" | "closed";
export type ThreadSpawnStatusFilter = ThreadSpawnEdgeStatus | "all";

export interface ThreadSpawnEdge {
  parentThreadId: ThreadId;
  childThreadId: ThreadId;
  status: ThreadSpawnEdgeStatus;
  createdAt: number;
  closedAt?: number;
  label?: string;
}

export interface ThreadSpawnEdgeInput {
  createdAt?: number;
  label?: string;
  status?: ThreadSpawnEdgeStatus;
}

export interface ThreadSpawnDescendant {
  threadId: ThreadId;
  parentThreadId: ThreadId;
  depth: number;
  edge: ThreadSpawnEdge;
}

export interface ListThreadSpawnDescendantsOptions {
  status?: ThreadSpawnStatusFilter;
}

interface AgentGraphFile {
  version: 1;
  edges: ThreadSpawnEdge[];
}

/**
 * Agent 线程拓扑图存储。
 *
 * 关联模块：
 * - sessionStore.ts: 负责单个线程 rollout，本模块只保存线程之间的父子关系。
 * - interaction/ipcAgentHandlers.ts: 对外暴露线程关系写入与查询入口。
 */
export class AgentGraphStore {
  private readonly graphPath: string;

  constructor(sessionsRoot = getDefaultSessionsRoot()) {
    this.graphPath = path.join(sessionsRoot, "thread-graph.json");
  }

  async upsertThreadSpawnEdge(
    parentThreadId: ThreadId,
    childThreadId: ThreadId,
    input: ThreadSpawnEdgeInput = {},
  ): Promise<ThreadSpawnEdge> {
    const graph = await this.loadGraph();
    const existing = graph.edges.find(
      (edge) => edge.parentThreadId === parentThreadId && edge.childThreadId === childThreadId,
    );

    if (existing) {
      existing.status = input.status ?? existing.status;
      existing.label = input.label ?? existing.label;
      if (existing.status === "open") {
        existing.closedAt = undefined;
      }
      await this.saveGraph(graph);
      return existing;
    }

    const edge: ThreadSpawnEdge = {
      parentThreadId,
      childThreadId,
      status: input.status ?? "open",
      createdAt: input.createdAt ?? Date.now(),
      label: input.label,
    };
    graph.edges.push(edge);
    await this.saveGraph(graph);
    return edge;
  }

  async closeThreadSpawnEdge(
    parentThreadId: ThreadId,
    childThreadId: ThreadId,
    closedAt = Date.now(),
  ): Promise<ThreadSpawnEdge | null> {
    const graph = await this.loadGraph();
    const edge = graph.edges.find(
      (item) => item.parentThreadId === parentThreadId && item.childThreadId === childThreadId,
    );
    if (!edge) return null;

    edge.status = "closed";
    edge.closedAt = closedAt;
    await this.saveGraph(graph);
    return edge;
  }

  async listThreadSpawnDescendants(
    parentThreadId: ThreadId,
    options: ListThreadSpawnDescendantsOptions = {},
  ): Promise<ThreadSpawnDescendant[]> {
    const graph = await this.loadGraph();
    const status = options.status ?? "all";
    const childrenByParent = new Map<ThreadId, ThreadSpawnEdge[]>();

    for (const edge of graph.edges) {
      if (status !== "all" && edge.status !== status) continue;
      const children = childrenByParent.get(edge.parentThreadId);
      if (children) {
        children.push(edge);
      } else {
        childrenByParent.set(edge.parentThreadId, [edge]);
      }
    }

    const descendants: ThreadSpawnDescendant[] = [];
    const visited = new Set<ThreadId>([parentThreadId]);
    const queue: Array<{ edge: ThreadSpawnEdge; depth: number }> =
      childrenByParent.get(parentThreadId)?.map((edge) => ({ edge, depth: 1 })) ?? [];

    while (queue.length > 0) {
      const current = queue.shift()!;
      if (visited.has(current.edge.childThreadId)) continue;

      visited.add(current.edge.childThreadId);
      descendants.push({
        threadId: current.edge.childThreadId,
        parentThreadId: current.edge.parentThreadId,
        depth: current.depth,
        edge: current.edge,
      });

      const children = childrenByParent.get(current.edge.childThreadId) ?? [];
      for (const child of children) {
        queue.push({ edge: child, depth: current.depth + 1 });
      }
    }

    return descendants;
  }

  private async loadGraph(): Promise<AgentGraphFile> {
    try {
      const content = await fs.promises.readFile(this.graphPath, "utf-8");
      const parsed = JSON.parse(content) as AgentGraphFile;
      return {
        version: 1,
        edges: Array.isArray(parsed.edges) ? parsed.edges : [],
      };
    } catch (err: any) {
      if (err.code === "ENOENT") return { version: 1, edges: [] };
      throw err;
    }
  }

  private async saveGraph(graph: AgentGraphFile): Promise<void> {
    await fs.promises.mkdir(path.dirname(this.graphPath), { recursive: true });
    await fs.promises.writeFile(this.graphPath, JSON.stringify(graph, null, 2), "utf-8");
  }
}
