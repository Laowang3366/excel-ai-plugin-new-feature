export function isComposerSubmitKey(key: string, shiftKey: boolean, isComposing = false): boolean {
  return key === "Enter" && !shiftKey && !isComposing;
}

export type ComposerPrimaryAction = "send" | "stop";

export function getComposerPrimaryAction(
  isStreaming: boolean,
  hasInput: boolean,
): ComposerPrimaryAction {
  return isStreaming && !hasInput ? "stop" : "send";
}
