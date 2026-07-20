export type AppTab = "chat" | "host" | "tools" | "providers";

const TABS: ReadonlySet<AppTab> = new Set(["chat", "host", "tools", "providers"]);

/**
 * Map task-pane URLSearchParams to the existing App tab.
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

  // Reject encoded path tricks / oversized tokens by normalizing once.
  let page = pageRaw;
  try {
    page = decodeURIComponent(pageRaw);
  } catch {
    return "chat";
  }
  page = page.trim().toLowerCase();

  if (page === "providers") return "providers";
  if (page === "host") return "host";
  if (page === "tools") return "tools";
  if (page === "chat") return "chat";

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

  if (TABS.has(page as AppTab)) return page as AppTab;
  return "chat";
}
