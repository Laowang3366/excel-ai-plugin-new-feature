/**
 * WPS JSA ribbon entry (browser + WPS host only).
 * Opens a real task pane via CreateTaskPane; never Node/Electron/COM.
 */
(function () {
  "use strict";

  var STORAGE_PANE_ID = "WenggeExcelAi.taskPaneId";
  var STORAGE_ROUTE = "WenggeExcelAi.taskPaneRoute";
  var ALLOWED_ROUTES = { chat: true, providers: true, host: true, tools: true };

  /** @type {unknown} */
  var ribbonUI = null;

  function resolveApi() {
    try {
      if (typeof wps !== "undefined" && wps) return wps;
    } catch (_e) {
      /* ignore */
    }
    if (typeof window !== "undefined") {
      if (window.wps) return window.wps;
      if (window.Wps) return window.Wps;
    }
    return null;
  }

  function normalizeRoute(route) {
    var key = String(route || "").toLowerCase();
    if (ALLOWED_ROUTES[key]) return key;
    return "chat";
  }

  /**
   * Derive task-pane URL from the currently loaded local index page.
   * Clears prior query/hash; only sets whitelist page= query.
   */
  function buildTaskPaneUrl(route) {
    var page = normalizeRoute(route);
    var href = "";
    try {
      href = String(window.location && window.location.href ? window.location.href : "");
    } catch (_e) {
      href = "";
    }
    if (!href) {
      throw new Error("missing location for task pane URL");
    }
    var base;
    try {
      var parsed = new URL(href);
      if (parsed.protocol !== "file:" && parsed.protocol !== "http:" && parsed.protocol !== "https:") {
        throw new Error("unsupported location protocol for task pane");
      }
      // Never invent external hosts: only reuse the page we were loaded from.
      parsed.hash = "";
      parsed.search = "";
      // Drop trailing empty search and set page only.
      var out = parsed.toString();
      // Ensure no leftover ? before adding ours
      if (out.charAt(out.length - 1) === "?") out = out.slice(0, -1);
      return out + "?page=" + encodeURIComponent(page);
    } catch (_err) {
      // Manual strip for odd runtimes without URL
      base = href.split("#")[0];
      var q = base.indexOf("?");
      if (q >= 0) base = base.slice(0, q);
      if (!base) throw new Error("cannot derive task pane base URL");
      if (/^https?:\/\/(?!localhost(?:[:/]|$)|127\.0\.0\.1(?:[:/]|$))/i.test(base)) {
        // Still allow same-page derivation even for non-local http (host may load file://).
        // Hardcoding remote CDN is avoided by never substituting base.
      }
      return base + "?page=" + encodeURIComponent(page);
    }
  }

  function paneIdOf(pane) {
    if (!pane || typeof pane !== "object") return null;
    if (pane.ID != null && pane.ID !== "") return pane.ID;
    if (pane.Id != null && pane.Id !== "") return pane.Id;
    if (pane.id != null && pane.id !== "") return pane.id;
    return null;
  }

  /** PluginStorage is stringly; host GetTaskPane may expect number IDs. */
  function coercePaneId(id) {
    if (id == null || id === "") return null;
    var asString = String(id);
    if (/^-?\d+$/.test(asString)) {
      var n = Number(asString);
      if (!isNaN(n)) return n;
    }
    return id;
  }

  function storageGet(api, key) {
    try {
      var store = api && api.PluginStorage;
      if (!store) return null;
      if (typeof store.getItem === "function") return store.getItem(key);
      if (typeof store.GetItem === "function") return store.GetItem(key);
    } catch (_e) {
      /* storage failure must not block open */
    }
    return null;
  }

  function storageSet(api, key, value) {
    try {
      var store = api && api.PluginStorage;
      if (!store) return false;
      if (typeof store.setItem === "function") {
        store.setItem(key, String(value));
        return true;
      }
      if (typeof store.SetItem === "function") {
        store.SetItem(key, String(value));
        return true;
      }
    } catch (_e) {
      /* ignore */
    }
    return false;
  }

  function createTaskPane(api, url) {
    if (api && typeof api.CreateTaskPane === "function") {
      return api.CreateTaskPane(url);
    }
    if (api && api.Application && typeof api.Application.CreateTaskPane === "function") {
      return api.Application.CreateTaskPane(url);
    }
    return null;
  }

  function getTaskPane(api, id) {
    try {
      if (api && typeof api.GetTaskPane === "function") {
        return api.GetTaskPane(id);
      }
      if (api && api.Application && typeof api.Application.GetTaskPane === "function") {
        return api.Application.GetTaskPane(id);
      }
    } catch (_e) {
      return null;
    }
    return null;
  }

  function setVisible(pane, visible) {
    if (!pane) return false;
    try {
      pane.Visible = visible;
      return true;
    } catch (_e) {
      return false;
    }
  }

  function fallbackOpen(url) {
    try {
      if (typeof window.open === "function") {
        var win = window.open(url);
        if (win) return true;
      }
    } catch (_e) {
      /* continue to alert */
    }
    try {
      if (typeof window.alert === "function") {
        window.alert(
          "无法打开文格 AI 任务窗格。请确认 WPS 已加载本插件且任务窗格 API 可用后，从功能区重试。",
        );
      }
    } catch (_e2) {
      /* ignore */
    }
    return false;
  }

  function openTaskPaneForRoute(route) {
    var page = normalizeRoute(route);
    var url;
    try {
      url = buildTaskPaneUrl(page);
    } catch (_e) {
      return fallbackOpen("index.html?page=" + encodeURIComponent(page));
    }

    // Refuse obviously non-derived remote hardcodes (defense in depth).
    if (/^https?:\/\//i.test(url)) {
      // Derived from current location only; still reject if base was rewritten to foreign host vs location.
      try {
        var locHost = "";
        var urlHost = "";
        try {
          locHost = new URL(String(window.location.href)).host;
          urlHost = new URL(url).host;
        } catch (_u) {
          /* ignore */
        }
        if (locHost && urlHost && locHost !== urlHost) {
          return fallbackOpen(url);
        }
      } catch (_c) {
        /* ignore */
      }
    }

    var api = resolveApi();
    if (!api) {
      return fallbackOpen(url);
    }

    var prevId = storageGet(api, STORAGE_PANE_ID);
    var prevRoute = storageGet(api, STORAGE_ROUTE);

    if (prevId != null && String(prevId) !== "" && String(prevRoute) === page) {
      var reused = getTaskPane(api, coercePaneId(prevId));
      if (reused) {
        if (setVisible(reused, true)) return true;
        // Visible failed → fallback (do not claim success via focus-only)
        return fallbackOpen(url);
      }
      // stale id — create fresh
    }

    if (prevId != null && String(prevId) !== "" && String(prevRoute) !== page) {
      var oldPane = getTaskPane(api, coercePaneId(prevId));
      setVisible(oldPane, false);
    }

    var pane = createTaskPane(api, url);
    if (!pane) {
      return fallbackOpen(url);
    }
    if (!setVisible(pane, true)) {
      return fallbackOpen(url);
    }

    var id = paneIdOf(pane);
    if (id != null) {
      storageSet(api, STORAGE_PANE_ID, id);
      storageSet(api, STORAGE_ROUTE, page);
    }
    return true;
  }

  window.WenggeExcelAiOnLoad = function (ui) {
    ribbonUI = ui || null;
  };

  window.WenggeExcelAiTabVisible = function () {
    return true;
  };

  window.WenggeExcelAiOpenChat = function () {
    return openTaskPaneForRoute("chat");
  };

  window.WenggeExcelAiOpenProviders = function () {
    return openTaskPaneForRoute("providers");
  };

  window.WenggeExcelAiOpenHost = function () {
    return openTaskPaneForRoute("host");
  };

  // Expose for tests only (idempotent, no auto-open side effects at load).
  window.__WenggeExcelAiEntryTest = {
    buildTaskPaneUrl: buildTaskPaneUrl,
    normalizeRoute: normalizeRoute,
    resolveApi: resolveApi,
    openTaskPaneForRoute: openTaskPaneForRoute,
    getRibbonUI: function () {
      return ribbonUI;
    },
  };
})();
