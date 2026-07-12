export function resolveFinalDownloadUrl(requestUrl: URL, responseUrl: string): URL {
  const finalUrl = responseUrl.trim();
  return finalUrl ? new URL(finalUrl, requestUrl) : requestUrl;
}
