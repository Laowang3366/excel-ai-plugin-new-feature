import { ipcMain } from "electron";
import {
  DEFAULT_RULES,
  setExtraWritableRoots,
  setUserRules,
  type PrefixRule,
} from "../agent/security/sandbox";
import {
  SandboxUserRulesInput,
  SandboxWritableRootsInput,
  validateInput,
} from "../shared/ipcSchemas";
import { getSettingsStore } from "./settingsManager";

export function registerSandboxIpcHandlers(): void {
  ipcMain.handle("sandbox:getConfig", () => {
    const store = getSettingsStore();
    return {
      defaultRules: getSandboxDefaultRulesForUI(),
      userRules: store.get("sandboxUserRules") as PrefixRule[] | undefined ?? [],
      extraWritableRoots: store.get("sandboxExtraWritableRoots") as string[] | undefined ?? [],
    };
  });

  ipcMain.handle("sandbox:setUserRules", (_event, rules: unknown) => {
    const validated = validateInput(SandboxUserRulesInput, rules);
    const normalized = normalizeUserRules(validated);
    if (normalized.error) {
      return { success: false, error: normalized.error };
    }
    getSettingsStore().set("sandboxUserRules", normalized.rules);
    applySandboxConfig();
    return { success: true };
  });

  ipcMain.handle("sandbox:setWritableRoots", (_event, roots: unknown) => {
    const clean = validateInput(SandboxWritableRootsInput, roots)
      .map((root) => root.trim())
      .filter(Boolean);
    getSettingsStore().set("sandboxExtraWritableRoots", clean);
    applySandboxConfig();
    return { success: true };
  });
}

function getSandboxDefaultRulesForUI(): PrefixRule[] {
  return JSON.parse(JSON.stringify(DEFAULT_RULES));
}

interface NormalizedRulesResult {
  rules: PrefixRule[];
  error?: string;
}

function normalizeUserRules(input: unknown[]): NormalizedRulesResult {
  const rules: PrefixRule[] = [];
  for (let i = 0; i < input.length; i++) {
    const raw = input[i];
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
      return { rules, error: `规则 #${i} 必须为对象` };
    }
    const r = raw as Record<string, unknown>;
    const pat = r.pattern;
    const firstFromPat = Array.isArray(pat) && pat.length > 0 ? pat[0] : undefined;
    const first = r.first ?? firstFromPat;
    if (typeof first !== "string" || first.length === 0) {
      return { rules, error: `规则 #${i} 缺少 first 或 first 非字符串` };
    }
    const decision = r.decision;
    if (decision !== "allow" && decision !== "prompt" && decision !== "forbidden") {
      return { rules, error: `规则 #${i} decision 必须为 allow/prompt/forbidden` };
    }
    const rest: PrefixRule["rest"] = [];
    const tailPattern: unknown[] = Array.isArray(pat) ? pat.slice(1) : Array.isArray(r.rest) ? r.rest as unknown[] : [];
    for (const tok of tailPattern) {
      if (Array.isArray(tok)) {
        if (tok.length === 0) return { rules, error: `规则 #${i} alts 不能为空` };
        rest.push(tok.length === 1
          ? { kind: "single", value: tok[0] as string }
          : { kind: "alts", values: tok as string[] });
      } else if (typeof tok === "string") {
        rest.push({ kind: "single", value: tok });
      } else {
        return { rules, error: `规则 #${i} pattern token 非法` };
      }
    }
    rules.push({
      first,
      rest,
      decision,
      justification: typeof r.justification === "string" ? r.justification : undefined,
    });
  }
  return { rules };
}

export function applySandboxConfig(): void {
  const store = getSettingsStore();
  const userRules = (store.get("sandboxUserRules") as PrefixRule[] | undefined) ?? [];
  const roots = (store.get("sandboxExtraWritableRoots") as string[] | undefined) ?? [];
  setUserRules(userRules ?? []);
  setExtraWritableRoots(roots ?? []);
}
