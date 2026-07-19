export interface RuntimeDateParts {
  CURRENT_DATE: string;
  CURRENT_TIME: string;
}

/** Same Asia/Shanghai Intl formatting as desktop buildRuntimePromptSection. */
export function formatRuntimeDateTime(now: Date = new Date()): RuntimeDateParts {
  const dateFormatter = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    weekday: "long",
  });
  const timeFormatter = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });
  return {
    CURRENT_DATE: dateFormatter.format(now),
    CURRENT_TIME: timeFormatter.format(now),
  };
}

/** Desktop uses: `- Office 应用连接状态：${status}` */
export function formatOfficeConnectionContext(status: string): string {
  return `- Office 应用连接状态：${status}`;
}
