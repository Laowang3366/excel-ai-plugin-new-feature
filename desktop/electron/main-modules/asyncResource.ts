export class AsyncResource<T> {
  private instance: T | null = null;
  private initializing: Promise<T> | null = null;
  private closing: Promise<void> | null = null;

  constructor(
    private readonly create: () => Promise<T>,
    private readonly dispose: (instance: T) => Promise<void>,
  ) {}

  async get(): Promise<T> {
    if (this.closing) await this.closing;
    if (this.instance) return this.instance;
    if (!this.initializing) {
      this.initializing = this.create()
        .then((instance) => {
          this.instance = instance;
          return instance;
        })
        .catch((error) => {
          this.initializing = null;
          throw error;
        });
    }
    return this.initializing;
  }

  async close(): Promise<void> {
    if (this.closing) return this.closing;
    if (!this.initializing && !this.instance) return;

    const pending = this.initializing;
    const active = this.instance;
    this.closing = (async () => {
      const instance = active ?? (await pending);
      if (instance) await this.dispose(instance);
    })().finally(() => {
      this.instance = null;
      this.initializing = null;
      this.closing = null;
    });
    return this.closing;
  }
}
