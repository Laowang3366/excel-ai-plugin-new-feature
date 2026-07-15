(function () {
  "use strict";

  var config = window.WENGGE_JSA_BRIDGE;
  if (!config) return;
  var baseUrl = "http://127.0.0.1:" + config.port;

  window.WenggeBridgeTabVisible = function () {
    return false;
  };

  function request(method, url, body, callback) {
    var xhr =
      typeof WpsInvoke !== "undefined" && WpsInvoke.CreateXHR
        ? WpsInvoke.CreateXHR()
        : new XMLHttpRequest();
    var completed = false;
    function finish(status, text) {
      if (completed) return;
      completed = true;
      callback(status, text || "");
    }
    xhr.open(method, baseUrl + url, true);
    xhr.timeout = 3000;
    xhr.setRequestHeader("X-Wengge-Token", config.token);
    if (body !== null) xhr.setRequestHeader("Content-Type", "application/json");
    xhr.onreadystatechange = function () {
      if (xhr.readyState === 4) finish(xhr.status, xhr.responseText);
    };
    xhr.onerror = function () {
      finish(0, "");
    };
    xhr.ontimeout = function () {
      finish(0, "");
    };
    try {
      xhr.send(body === null ? null : JSON.stringify(body));
    } catch (_) {
      finish(0, "");
    }
  }

  function getCodeModule() {
    var app = window.Application;
    if (!app || !app.ActiveWorkbook) throw new Error("当前没有活动的 WPS 工作簿");
    var component = app.JSIDE.SelectedJSComponent;
    if (!component || !component.CodeModule) {
      throw new Error("无法访问当前 JSA 组件，请在宏安全性中信任对 wpsjs 项目的访问");
    }
    return { app: app, component: component, module: component.CodeModule };
  }

  function readSource(module) {
    var count = module.CountOfLines;
    return count > 0 ? module.Lines(1, count) : "";
  }

  function normalize(source) {
    return String(source || "")
      .replace(/\r\n?/g, "\n")
      .replace(/\n+$/g, "")
      .replace(/^\s+/, "");
  }

  function hasEntryPoint(source, entryPoint) {
    if (!entryPoint) return true;
    var name = String(entryPoint).split(".").pop();
    var escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    var pattern =
      "(?:^|\\n)\\s*(?:export\\s+)?(?:async\\s+)?function\\s+" +
      escaped +
      "\\s*\\(|(?:^|\\n)\\s*(?:export\\s+)?(?:const|let|var)\\s+" +
      escaped +
      "\\s*=";
    return new RegExp(pattern, "m").test(source);
  }

  function writeCode(data) {
    var target = getCodeModule();
    var oldSource = readSource(target.module);
    try {
      if (target.module.CountOfLines > 0) target.module.DeleteLines(1, target.module.CountOfLines);
      if (data.code) target.module.AddFromString(data.code);
      var source = readSource(target.module);
      if (normalize(source) !== normalize(data.code)) throw new Error("JSA 源码回读不一致");
      if (!hasEntryPoint(source, data.entryPoint))
        throw new Error("写入后找不到入口函数: " + data.entryPoint);
      if (data.save) target.app.ActiveWorkbook.Save();
      return {
        componentName: target.component.Name || "",
        lineCount: target.module.CountOfLines,
        source: source,
        entryPointVerified: true,
        saved: data.save === true,
        workbookName: target.app.ActiveWorkbook.Name || "",
      };
    } catch (error) {
      try {
        if (target.module.CountOfLines > 0)
          target.module.DeleteLines(1, target.module.CountOfLines);
        if (oldSource) target.module.AddFromString(oldSource);
      } catch (_) {
        // Keep the original write error when rollback also fails.
      }
      throw error;
    }
  }

  function execute(command) {
    if (command.type === "detect") {
      var target = getCodeModule();
      return {
        componentName: target.component.Name || "",
        workbookName: target.app.ActiveWorkbook.Name || "",
      };
    }
    if (command.type === "write") return writeCode(command.data || {});
    throw new Error("未知 JSA 桥接命令: " + command.type);
  }

  function poll() {
    request("GET", "/command", null, function (status, text) {
      if (status === 200 && text) {
        var command;
        var response;
        try {
          command = JSON.parse(text);
          response = { id: command.id, ok: true, result: execute(command) };
        } catch (error) {
          response = {
            id: (command && command.id) || "",
            ok: false,
            error: error.message || String(error),
          };
        }
        request("POST", "/response", response, function () {});
      }
      window.setTimeout(poll, 400);
    });
  }

  poll();
})();
