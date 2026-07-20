/**
 * WPS JSA ribbon entry (browser + WPS host only).
 * Opens a real task pane via CreateTaskPane; never Node/Electron/COM.
 */
(function () {
  "use strict";

  var STORAGE_PANE_ID = "WenggeExcelAi.taskPaneId";
  var STORAGE_ROUTE = "WenggeExcelAi.taskPaneRoute";
  var ALLOWED_ROUTES = { chat: true, providers: true, host: true, tools: true };
  var ICON_32 = "assets/icon-32.png";
  var ICON_16 = "assets/icon-16.png";

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

  function hasControlChars(value) {
    return /[\u0000-\u001f\u007f]/.test(String(value));
  }

  function isAllowedProtocol(protocol) {
    return protocol === "file:" || protocol === "http:" || protocol === "https:";
  }

  function readLocationHref() {
    try {
      if (!window || !window.location) {
        throw new Error("missing location");
      }
      var href = window.location.href;
      if (href == null || href === "") {
        throw new Error("missing location href");
      }
      return String(href);
    } catch (error) {
      throw new Error(
        error instanceof Error ? error.message : "location unavailable",
      );
    }
  }

  function rejectForbiddenHref(href) {
    if (!href || hasControlChars(href)) {
      throw new Error("missing or unsafe location for task pane URL");
    }
    if (/^\s*(javascript|data|blob|about|vbscript):/i.test(href)) {
      throw new Error("forbidden location protocol for task pane");
    }
  }

  /**
   * Manual base strip only when URL constructor is unavailable.
   * Still requires an explicit file/http/https scheme whitelist.
   */
  function buildTaskPaneUrlManual(href, page) {
    rejectForbiddenHref(href);
    var scheme = href.match(/^(file:|https?:)/i);
    if (!scheme) {
      throw new Error("unsupported location protocol for task pane");
    }
    var base = href.split("#")[0];
    var q = base.indexOf("?");
    if (q >= 0) base = base.slice(0, q);
    if (!base || hasControlChars(base)) {
      throw new Error("cannot derive task pane base URL");
    }
    if (!/^(file:|https?:)/i.test(base)) {
      throw new Error("unsupported location protocol for task pane");
    }
    return base + "?page=" + encodeURIComponent(page);
  }

  /**
   * Derive task-pane URL from the currently loaded local index page.
   * Clears prior query/hash; only sets whitelist page= query.
   * Fail closed for javascript/data/blob/about and control characters.
   */
  function buildTaskPaneUrl(route) {
    var page = normalizeRoute(route);
    var href = readLocationHref();
    rejectForbiddenHref(href);

    if (typeof URL !== "function") {
      return buildTaskPaneUrlManual(href, page);
    }

    var loc;
    try {
      loc = new URL(href);
    } catch (_parseErr) {
      throw new Error("cannot parse location for task pane URL");
    }

    if (!isAllowedProtocol(loc.protocol)) {
      throw new Error("unsupported location protocol for task pane");
    }

    loc.hash = "";
    loc.search = "";
    var base = loc.toString();
    if (base.charAt(base.length - 1) === "?") {
      base = base.slice(0, -1);
    }
    var out = base + "?page=" + encodeURIComponent(page);

    var derived;
    try {
      derived = new URL(out);
    } catch (_outErr) {
      throw new Error("derived task pane URL is invalid");
    }
    if (!isAllowedProtocol(derived.protocol)) {
      throw new Error("derived task pane URL has forbidden protocol");
    }
    if (loc.protocol === "file:") {
      if (derived.protocol !== "file:" || derived.pathname !== loc.pathname) {
        throw new Error("derived task pane URL left the local index path");
      }
    } else if (derived.origin !== loc.origin) {
      throw new Error("derived task pane URL left the current origin");
    }
    return out;
  }

  function isSafeOpenUrl(url) {
    if (typeof url !== "string" || !url || hasControlChars(url)) return false;
    if (/^\s*(javascript|data|blob|about|vbscript):/i.test(url)) return false;
    try {
      var href = readLocationHref();
      rejectForbiddenHref(href);
      if (typeof URL === "function") {
        var loc = new URL(href);
        var candidate = new URL(url);
        if (!isAllowedProtocol(loc.protocol) || !isAllowedProtocol(candidate.protocol)) {
          return false;
        }
        if (loc.protocol === "file:") {
          return candidate.protocol === "file:" && candidate.pathname === loc.pathname;
        }
        return candidate.origin === loc.origin;
      }
      return /^(file:|https?:)/i.test(url);
    } catch (_e) {
      return false;
    }
  }

  function paneIdOf(pane) {
    if (!pane || typeof pane !== "object") return null;
    try {
      if (pane.ID != null && pane.ID !== "") return pane.ID;
    } catch (_e1) {
      /* ignore getter throw */
    }
    try {
      if (pane.Id != null && pane.Id !== "") return pane.Id;
    } catch (_e2) {
      /* ignore */
    }
    try {
      if (pane.id != null && pane.id !== "") return pane.id;
    } catch (_e3) {
      /* ignore */
    }
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
    if (!api || !url) return null;
    try {
      if (typeof api.CreateTaskPane === "function") {
        return api.CreateTaskPane(url);
      }
    } catch (_e1) {
      /* try Application fallback */
    }
    try {
      if (api.Application && typeof api.Application.CreateTaskPane === "function") {
        return api.Application.CreateTaskPane(url);
      }
    } catch (_e2) {
      return null;
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

  function alertUnavailable() {
    try {
      if (typeof window.alert === "function") {
        window.alert(
          "无法打开文格 AI 任务窗格。请确认 WPS 已加载本插件且任务窗格 API 可用后，从功能区重试。",
        );
      }
    } catch (_e) {
      /* ignore */
    }
  }

  /**
   * @param {string|null|undefined} url
   * Safe URLs may window.open; unsafe/null only alert and return false.
   */
  function fallbackOpen(url) {
    if (url && isSafeOpenUrl(url)) {
      try {
        if (typeof window.open === "function") {
          var win = window.open(url);
          if (win) return true;
        }
      } catch (_e) {
        /* continue to alert */
      }
    }
    alertUnavailable();
    return false;
  }

  function openTaskPaneForRoute(route) {
    var page = normalizeRoute(route);
    var url;
    try {
      url = buildTaskPaneUrl(page);
    } catch (_e) {
      return fallbackOpen(null);
    }

    if (!isSafeOpenUrl(url)) {
      return fallbackOpen(null);
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
        return fallbackOpen(url);
      }
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

    // Pane is already visible; ID getter failures must not undo success.
    try {
      var id = paneIdOf(pane);
      if (id != null) {
        storageSet(api, STORAGE_PANE_ID, id);
        storageSet(api, STORAGE_ROUTE, page);
      }
    } catch (_idErr) {
      /* ignore */
    }
    return true;
  }

  function controlIdOf(control) {
    if (control == null) return "";
    if (typeof control === "string") return control;
    if (typeof control !== "object") return "";
    try {
      if (control.Id != null && control.Id !== "") return String(control.Id);
    } catch (_e1) {
      /* ignore */
    }
    try {
      if (control.ID != null && control.ID !== "") return String(control.ID);
    } catch (_e2) {
      /* ignore */
    }
    try {
      if (control.id != null && control.id !== "") return String(control.id);
    } catch (_e3) {
      /* ignore */
    }
    return "";
  }

  /**
   * Ribbon getImage callback (host-proven pattern).
   * Returns only fixed package-relative PNG paths — never user/path/URL input.
   */
  window.WenggeExcelAiGetImage = function (control) {
    var id = controlIdOf(control);
    if (id === "wenggeExcelAiOpenHostButton") return ICON_16;
    if (id === "wenggeExcelAiOpenChatButton") return ICON_32;
    if (id === "wenggeExcelAiOpenProvidersButton") return ICON_32;
    return ICON_32;
  };

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
    getImage: window.WenggeExcelAiGetImage,
    controlIdOf: controlIdOf,
    getRibbonUI: function () {
      return ribbonUI;
    },
  };
})();
