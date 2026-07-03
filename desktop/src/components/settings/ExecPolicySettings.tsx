/**
 * 安全策略设置 — 命令策略（execpolicy）与可写根管理
 *
 * 对应 docs/sandbox-implementation-plan.md 阶段 1：
 * - 展示内置默认规则（只读）
 * - 让用户追加自定义规则（前缀 token + decision + justification）
 * - 管理额外可写根目录
 *
 * 这里只做最小可用 UI：列规则、增加/删除用户规则、可写根增删，
 * 复杂的可视化规则编辑（互斥 token 编辑器等）后续再加。
 */

import React, { useEffect, useState } from "react";
import { ShieldAlert, Plus, Trash2, Check } from "../common/IconMap";
import { useSettingsStore } from "../../store/settingsStore";
import { ipcApi } from "../../services/ipcApi";
import type { SandboxPrefixRule } from "../../electronApi";

const TEXT = {
  "zh-CN": {
    title: "命令安全策略",
    desc: "shell.execute 工具的命令会先经策略评估：forbidden 拒绝、prompt 强制审批、allow 放行。规则按命令前缀 token 匹配，多规则命中取最严。",
    defaultRulesTitle: "内置默认规则（只读）",
    userRulesTitle: "自定义规则",
    userRulesHint: "在默认规则之上追加；first 必填，decision 选 allow / prompt / forbidden。",
    addUserRule: "添加规则",
    save: "保存",
    saved: "已保存",
    saveFail: "保存失败",
    add: "添加",
    cancel: "取消",
    del: "删除",
    fieldFirst: "命令名（首 token）",
    fieldDecision: "决策",
    fieldJustification: "理由（可选）",
    decisionAllow: "放行",
    decisionPrompt: "询问",
    decisionForbidden: "拒绝",
    rootsTitle: "可写根目录",
    rootsHint: "shell.execute 的工作目录只能在以下根（含子目录）内：默认临时目录、桌面、文档、下载 + 你自定义的下列根。越界自动重定向到临时目录。",
    addRoot: "添加根目录",
    newRootPlaceholder: "C:\\Projects 或 /home/me/repos",
    loading: "加载中...",
  },
  "en-US": {
    title: "Command Security Policy",
    desc: "Commands from shell.execute are evaluated first: forbidden rejects, prompt forces approval, allow passes. Rules match by command prefix tokens; multiple matches take the strictest.",
    defaultRulesTitle: "Built-in default rules (read-only)",
    userRulesTitle: "Custom rules",
    userRulesHint: "Appended on top of the defaults; first is required, decision ∈ allow / prompt / forbidden.",
    addUserRule: "Add rule",
    save: "Save",
    saved: "Saved",
    saveFail: "Save failed",
    add: "Add",
    cancel: "Cancel",
    del: "Delete",
    fieldFirst: "Command name (first token)",
    fieldDecision: "Decision",
    fieldJustification: "Reason (optional)",
    decisionAllow: "Allow",
    decisionPrompt: "Prompt",
    decisionForbidden: "Forbid",
    rootsTitle: "Writable roots",
    rootsHint: "shell.execute cwd must be inside one of: default tmp, Desktop, Documents, Downloads + your custom roots below. Out-of-bounds cwd is auto-redirected to tmp.",
    addRoot: "Add root",
    newRootPlaceholder: "C:\\Projects or /home/me/repos",
    loading: "Loading...",
  },
} as const;

const DECISION_LABEL = {
  "zh-CN": { allow: "放行", prompt: "询问", forbidden: "拒绝" },
  "en-US": { allow: "Allow", prompt: "Prompt", forbidden: "Forbid" },
} as const;

export const ExecPolicySettings: React.FC = () => {
  const { language } = useSettingsStore();
  const t = TEXT[language];
  const dl = DECISION_LABEL[language];

  const [defaultRules, setDefaultRules] = useState<SandboxPrefixRule[]>([]);
  const [userRules, setUserRules] = useState<SandboxPrefixRule[]>([]);
  const [writableRoots, setWritableRoots] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);

  // 新规则表单
  const [newFirst, setNewFirst] = useState("");
  const [newDecision, setNewDecision] = useState<SandboxPrefixRule["decision"]>("prompt");
  const [newJustification, setNewJustification] = useState("");

  // 新根
  const [newRoot, setNewRoot] = useState("");

  // 保存反馈
  const [saveMsg, setSaveMsg] = useState<{ ok: boolean; text: string } | null>(null);

  const reload = async () => {
    setLoading(true);
    try {
      const cfg = await ipcApi.sandbox.getConfig();
      setDefaultRules(cfg.defaultRules ?? []);
      setUserRules(cfg.userRules ?? []);
      setWritableRoots(cfg.extraWritableRoots ?? []);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    reload();
  }, []);

  const addUserRule = () => {
    const first = newFirst.trim();
    if (!first) return;
    setUserRules([
      ...userRules,
      { first, rest: [], decision: newDecision, justification: newJustification || undefined },
    ]);
    setNewFirst("");
    setNewJustification("");
  };

  const removeUserRule = (idx: number) => {
    setUserRules(userRules.filter((_, i) => i !== idx));
  };

  const addRoot = () => {
    const r = newRoot.trim();
    if (!r) return;
    setWritableRoots([...writableRoots, r]);
    setNewRoot("");
  };

  const removeRoot = (idx: number) => {
    setWritableRoots(writableRoots.filter((_, i) => i !== idx));
  };

  const saveAll = async () => {
    try {
      const r1 = await ipcApi.sandbox.setUserRules(userRules);
      if (!r1.success) {
        setSaveMsg({ ok: false, text: r1.error || t.saveFail });
        return;
      }
      const r2 = await ipcApi.sandbox.setWritableRoots(writableRoots);
      if (!r2.success) {
        setSaveMsg({ ok: false, text: r2.error || t.saveFail });
        return;
      }
      setSaveMsg({ ok: true, text: t.saved });
      setTimeout(() => setSaveMsg(null), 2000);
    } catch (err: any) {
      setSaveMsg({ ok: false, text: err?.message || t.saveFail });
    }
  };

  const renderRuleRow = (r: SandboxPrefixRule, idx?: number, onDelete?: () => void) => {
    const rest = r.rest.map((tok) =>
      tok.kind === "single" ? tok.value : `[${tok.values.join("|")}]`
    );
    const pattern = [r.first, ...rest].join(" ");
    return (
      <div className="exec-policy-row" key={idx ?? pattern + r.decision}>
        <code className="exec-policy-pattern">{pattern}</code>
        <span className={`exec-policy-decision exec-policy-decision-${r.decision}`}>
          {dl[r.decision]}
        </span>
        {r.justification && <span className="exec-policy-just">{r.justification}</span>}
        {onDelete && (
          <button className="exec-policy-del" onClick={onDelete} title={TEXT[language].del}>
            <Trash2 size={13} />
          </button>
        )}
      </div>
    );
  };

  if (loading) {
    return <div className="section-loading">{TEXT[language].loading}</div>;
  }

  return (
    <div className="exec-policy-settings">
      <h2 className="section-title"><ShieldAlert size={18} /> {t.title}</h2>
      <p className="section-desc">{t.desc}</p>

      {/* 默认规则 */}
      <div className="exec-policy-block">
        <h3 className="exec-policy-subtitle">{t.defaultRulesTitle}</h3>
        <div className="exec-policy-list">
          {defaultRules.map((r) => renderRuleRow(r))}
        </div>
      </div>

      {/* 用户规则 */}
      <div className="exec-policy-block">
        <h3 className="exec-policy-subtitle">{t.userRulesTitle}</h3>
        <p className="exec-policy-hint">{t.userRulesHint}</p>

        <div className="exec-policy-add">
          <input
            className="exec-policy-input"
            placeholder={t.fieldFirst}
            value={newFirst}
            onChange={(e) => setNewFirst(e.target.value)}
          />
          <select
            className="exec-policy-select"
            value={newDecision}
            onChange={(e) => setNewDecision(e.target.value as SandboxPrefixRule["decision"])}
          >
            <option value="allow">{t.decisionAllow}</option>
            <option value="prompt">{t.decisionPrompt}</option>
            <option value="forbidden">{t.decisionForbidden}</option>
          </select>
          <input
            className="exec-policy-input exec-policy-input-just"
            placeholder={t.fieldJustification}
            value={newJustification}
            onChange={(e) => setNewJustification(e.target.value)}
          />
          <button className="exec-policy-add-btn" onClick={addUserRule}>
            <Plus size={13} /> {t.addUserRule}
          </button>
        </div>

        <div className="exec-policy-list">
          {userRules.length === 0 ? (
            <div className="exec-policy-empty">—</div>
          ) : (
            userRules.map((r, i) => renderRuleRow(r, i, () => removeUserRule(i)))
          )}
        </div>
      </div>

      {/* 可写根 */}
      <div className="exec-policy-block">
        <h3 className="exec-policy-subtitle">{t.rootsTitle}</h3>
        <p className="exec-policy-hint">{t.rootsHint}</p>
        <div className="exec-policy-roots">
          {writableRoots.map((r, i) => (
            <div className="exec-policy-root-row" key={r + i}>
              <code>{r}</code>
              <button className="exec-policy-del" onClick={() => removeRoot(i)} title={TEXT[language].del}>
                <Trash2 size={13} />
              </button>
            </div>
          ))}
          <div className="exec-policy-add">
            <input
              className="exec-policy-input exec-policy-input-root"
              placeholder={t.newRootPlaceholder}
              value={newRoot}
              onChange={(e) => setNewRoot(e.target.value)}
            />
            <button className="exec-policy-add-btn" onClick={addRoot}>
              <Plus size={13} /> {t.addRoot}
            </button>
          </div>
        </div>
      </div>

      {/* 保存条 */}
      <div className="exec-policy-save-bar">
        <button className="exec-policy-save-btn" onClick={saveAll}>
          <Check size={13} /> {t.save}
        </button>
        {saveMsg && (
          <span className={`exec-policy-save-msg ${saveMsg.ok ? "ok" : "err"}`}>
            {saveMsg.text}
          </span>
        )}
      </div>
    </div>
  );
};