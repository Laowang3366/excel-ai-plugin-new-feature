import type { HostAdapter } from "../host/hostAdapter";

export interface ChatPromptRuntimeContext {
  officeConnectionStatus: string;
  dynamicArrayEnabled: boolean;
}

export async function resolveChatPromptRuntimeContext(
  host: HostAdapter,
): Promise<ChatPromptRuntimeContext> {
  let dynamicArrayEnabled = false;
  try {
    dynamicArrayEnabled =
      host.getRuntimeCapabilities().dynamicArrayFunctionsEnabled === true;
  } catch {
    // A capability probe must never prevent chat from starting.
  }

  let officeConnectionStatus = `unavailable (${host.kind})`;
  try {
    const status = await host.getStatus();
    if (status.ok) {
      const kind = status.data.kind === "unknown" ? host.kind : status.data.kind;
      officeConnectionStatus = `${
        status.data.connected ? "connected" : "disconnected"
      } (${kind})`;
    }
  } catch {
    // Host status is advisory prompt context, not a chat precondition.
  }

  return { officeConnectionStatus, dynamicArrayEnabled };
}
