export function isAllowedWindowNavigation(targetUrl: string, appUrl: string): boolean {
  try {
    const target = new URL(targetUrl);
    const allowed = new URL(appUrl);
    if (target.protocol !== allowed.protocol || target.origin !== allowed.origin) return false;
    return target.pathname === allowed.pathname;
  } catch {
    return false;
  }
}
