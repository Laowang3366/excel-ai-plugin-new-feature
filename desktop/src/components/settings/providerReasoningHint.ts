import type { AppLanguage } from "../../store/settingsStore";
import type { ReasoningMode } from "../../electronApi";
import { formatReasoningOptionLabels } from "../../utils/reasoningSupport";

export function buildReasoningAutoHint(
  reasoningOptionValues: ReasoningMode[],
  language: AppLanguage
): string {
  const labels = formatReasoningOptionLabels(reasoningOptionValues, language);
  return language === "zh-CN"
    ? `已根据当前供应商/API/模型自动适配：${labels}`
    : `Automatically adapted for this provider/API/model: ${labels}`;
}
