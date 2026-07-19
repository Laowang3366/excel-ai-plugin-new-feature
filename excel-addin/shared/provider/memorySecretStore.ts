import type { ProviderSecretStore } from "./types";

/** In-memory API key store. Never write keys to localStorage. */
export class MemorySecretStore implements ProviderSecretStore {
  private readonly keys = new Map<string, string>();

  get(id: string): string | undefined {
    return this.keys.get(id);
  }

  set(id: string, apiKey: string): void {
    this.keys.set(id, apiKey);
  }

  delete(id: string): void {
    this.keys.delete(id);
  }

  clear(): void {
    this.keys.clear();
  }
}
