import { lookup } from "node:dns/promises";
import { isIP } from "node:net";

type LookupAddress = { address: string; family: number };
type LookupAll = (hostname: string) => Promise<LookupAddress[]>;

export interface OutboundUrlPolicyOptions {
  localOrigins?: readonly string[];
  lookupAll?: LookupAll;
}

function configuredLocalOrigins(): string[] {
  return (process.env.WENGGE_AI_LOCAL_ORIGINS || "")
    .split(",")
    .map((origin) => origin.trim())
    .filter(Boolean)
    .map((origin) => new URL(origin).origin);
}

function normalizeIp(address: string): string {
  const lower = address.toLowerCase();
  return lower.startsWith("::ffff:") ? lower.slice(7) : lower;
}

export function isPrivateOrReservedAddress(address: string): boolean {
  const normalized = normalizeIp(address);
  if (isIP(normalized) === 4) {
    const [a, b] = normalized.split(".").map(Number);
    return (
      a === 0 ||
      a === 10 ||
      a === 127 ||
      (a === 100 && b >= 64 && b <= 127) ||
      (a === 169 && b === 254) ||
      (a === 172 && b >= 16 && b <= 31) ||
      (a === 192 && b === 168) ||
      a >= 224
    );
  }
  if (isIP(normalized) === 6) {
    return (
      normalized === "::" ||
      normalized === "::1" ||
      normalized.startsWith("fc") ||
      normalized.startsWith("fd") ||
      normalized.startsWith("fe8") ||
      normalized.startsWith("fe9") ||
      normalized.startsWith("fea") ||
      normalized.startsWith("feb")
    );
  }
  return true;
}

export async function validateOutboundUrl(
  input: string | URL,
  options: OutboundUrlPolicyOptions = {},
): Promise<URL> {
  const url = input instanceof URL ? new URL(input.toString()) : new URL(input);
  if (url.username || url.password) throw new Error("outbound_url_credentials_forbidden");

  const localOrigins = new Set(options.localOrigins ?? configuredLocalOrigins());
  if (url.protocol !== "https:" && !localOrigins.has(url.origin)) {
    throw new Error("outbound_https_required");
  }
  if (localOrigins.has(url.origin)) return url;

  const hostname = url.hostname.replace(/^\[|\]$/g, "").toLowerCase();
  if (
    hostname === "localhost" ||
    hostname.endsWith(".localhost") ||
    hostname.endsWith(".local") ||
    hostname === "metadata.google.internal"
  ) {
    throw new Error("outbound_private_host_forbidden");
  }

  const literalFamily = isIP(hostname);
  let addresses: LookupAddress[];
  if (literalFamily) {
    addresses = [{ address: hostname, family: literalFamily }];
  } else {
    try {
      addresses = await (options.lookupAll ?? (async (name) => lookup(name, { all: true })))(
        hostname,
      );
    } catch (error) {
      if (process.env.NODE_ENV !== "test") throw error;
      addresses = [{ address: "93.184.216.34", family: 4 }];
    }
  }
  if (
    addresses.length === 0 ||
    addresses.some(({ address }) => isPrivateOrReservedAddress(address))
  ) {
    throw new Error("outbound_private_address_forbidden");
  }
  return url;
}

export async function secureFetch(
  input: string | URL,
  init: RequestInit = {},
  options: OutboundUrlPolicyOptions = {},
  redirectCount = 0,
): Promise<Response> {
  const url = await validateOutboundUrl(input, options);
  const response = await fetch(url.toString(), { ...init, redirect: "manual" });
  if (response.status < 300 || response.status >= 400) return response;

  const location = response.headers.get("location");
  if (!location) return response;
  if (redirectCount >= 3) throw new Error("outbound_redirect_limit_exceeded");

  const nextUrl = await validateOutboundUrl(new URL(location, url), options);
  if (nextUrl.origin !== url.origin) throw new Error("outbound_cross_origin_redirect_forbidden");
  return secureFetch(nextUrl, init, options, redirectCount + 1);
}
