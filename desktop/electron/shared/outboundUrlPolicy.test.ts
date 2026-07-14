import { afterEach, describe, expect, it, vi } from "vitest";

import { isPrivateOrReservedAddress, secureFetch, validateOutboundUrl } from "./outboundUrlPolicy";

const publicLookup = async () => [{ address: "93.184.216.34", family: 4 }];

describe("outbound URL policy", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });
  it("allows public HTTPS endpoints", async () => {
    await expect(validateOutboundUrl("https://api.example.com/v1", {
      lookupAll: publicLookup,
    })).resolves.toMatchObject({ protocol: "https:", hostname: "api.example.com" });
  });

  it("rejects credentials, private hosts, metadata and encoded IP forms", async () => {
    await expect(validateOutboundUrl("https://user:pass@example.com", {
      lookupAll: publicLookup,
    })).rejects.toThrow("outbound_url_credentials_forbidden");
    await expect(validateOutboundUrl("https://127.0.0.1")).rejects.toThrow("outbound_private_address_forbidden");
    await expect(validateOutboundUrl("https://2130706433")).rejects.toThrow("outbound_private_address_forbidden");
    await expect(validateOutboundUrl("https://[::1]")).rejects.toThrow("outbound_private_address_forbidden");
    await expect(validateOutboundUrl("https://metadata.google.internal")).rejects.toThrow("outbound_private_host_forbidden");
  });

  it("checks every DNS answer and allows only explicitly configured local origins", async () => {
    await expect(validateOutboundUrl("https://rebind.example", {
      lookupAll: async () => [
        { address: "93.184.216.34", family: 4 },
        { address: "10.0.0.5", family: 4 },
      ],
    })).rejects.toThrow("outbound_private_address_forbidden");
    await expect(validateOutboundUrl("http://127.0.0.1:11434/v1", {
      localOrigins: ["http://127.0.0.1:11434"],
    })).resolves.toMatchObject({ origin: "http://127.0.0.1:11434" });
    await expect(validateOutboundUrl("http://127.0.0.1:11435/v1", {
      localOrigins: ["http://127.0.0.1:11434"],
    })).rejects.toThrow("outbound_https_required");
  });

  it("classifies private and reserved IP ranges", () => {
    expect(isPrivateOrReservedAddress("10.1.2.3")).toBe(true);
    expect(isPrivateOrReservedAddress("169.254.169.254")).toBe(true);
    expect(isPrivateOrReservedAddress("::ffff:192.168.1.1")).toBe(true);
    expect(isPrivateOrReservedAddress("8.8.8.8")).toBe(false);
    expect(isPrivateOrReservedAddress("2606:4700:4700::1111")).toBe(false);
  });

  it("does not forward credentials through redirects", async () => {
    const fetchMock = vi.fn(async () => new Response(null, {
      status: 302,
      headers: { location: "https://127.0.0.1/steal" },
    }));
    vi.stubGlobal("fetch", fetchMock);

    await expect(secureFetch("https://api.example.com/v1", {
      headers: { Authorization: "Bearer canary" },
    }, { lookupAll: publicLookup })).rejects.toThrow("outbound_private_address_forbidden");
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});
