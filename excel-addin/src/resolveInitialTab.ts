export type AppTab =
  | "chat"
  | "host"
  | "tools"
  | "providers"
  | "formula"
  | "clean"
  | "chart"
  | "report"
  | "ocr";

const TABS: ReadonlySet<AppTab> = new Set([
  "chat",
  "host",
  "tools",
  "providers",
  "formula",
  "clean",
  "chart",
  "report",
  "ocr",
]);

const PAGE_ALIASES: Record<string, AppTab> = {
  chat: "chat",
  host: "host",
  tools: "tools",
  providers: "providers",
  formula: "formula",
  formulas: "formula",
  clean: "clean",
  cleaning: "clean",
  chart: "chart",
  charts: "chart",
  report: "report",
  reports: "report",
  ocr: "ocr",
};

/**
 * Map task-pane URLSearchParams to the App tab.
 * Unknown / malicious values fall back to chat. Does not write to the DOM.
 */
export function resolveInitialTabFromSearch(
  search: string | URLSearchParams | null | undefined,
): AppTab {
  let params: URLSearchParams;
  if (search instanceof URLSearchParams) {
    params = search;
  } else if (typeof search === "string") {
    const raw = search.startsWith("?") ? search.slice(1) : search;
    params = new URLSearchParams(raw);
  } else {
    return "chat";
  }

  const pageRaw = params.get("page");
  if (pageRaw == null || pageRaw === "") return "chat";

  let page = pageRaw;
  try {
    page = decodeURIComponent(pageRaw);
  } catch {
    return "chat";
  }
  page = page.trim().toLowerCase();

  // Legacy deep-link: page=settings&section=model → providers
  if (page === "settings") {
    const sectionRaw = params.get("section") ?? "";
    let section = sectionRaw;
    try {
      section = decodeURIComponent(sectionRaw);
    } catch {
      return "chat";
    }
    if (section.trim().toLowerCase() === "model") return "providers";
  }

  const aliased = PAGE_ALIASES[page];
  if (aliased) return aliased;
  if (TABS.has(page as AppTab)) return page as AppTab;
  return "chat";
}
