export function resolveDevHttpsOptions(): Promise<{
  cert: Buffer;
  key: Buffer;
  ca?: Buffer;
}>;
