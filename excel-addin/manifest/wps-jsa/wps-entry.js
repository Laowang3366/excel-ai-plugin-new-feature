(function () {
  "use strict";

  window.WenggeExcelAiTabVisible = function () {
    return true;
  };

  window.WenggeExcelAiOpenTaskPane = function () {
    if (typeof window.focus === "function") window.focus();
    return true;
  };
})();
