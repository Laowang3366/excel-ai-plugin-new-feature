import {
  createContext,
  useContext,
  type ReactNode,
} from "react";
import type { ChatController, ChatControllerDeps } from "@shared/agentChat";
import type { HostAdapter } from "@shared/host";
import type { ProviderStore } from "@shared/provider";
import {
  useChatController,
  type ChatSendOutcome,
  type ChatViewState,
} from "./useChatController";
import type { AgentContentPart } from "@shared/agent";
import type { ChatToolExecuteResult } from "@shared/agentChat";

export interface ChatSessionValue {
  view: ChatViewState;
  send: (
    text: string,
    options?: { contentParts?: AgentContentPart[] },
  ) => Promise<ChatSendOutcome>;
  retry: (turnId: string) => Promise<ChatSendOutcome>;
  executeTool: (
    toolName: string,
    args: Record<string, unknown>,
    options?: { toolCallId?: string },
  ) => Promise<ChatToolExecuteResult>;
  stop: () => void;
  clear: () => void;
  approve: (requestId?: string) => boolean;
  reject: (requestId?: string) => boolean;
  controller: ChatController | null;
  adapter: HostAdapter | null;
  store: ProviderStore;
}

const ChatSessionContext = createContext<ChatSessionValue | null>(null);

export function ChatSessionProvider({
  store,
  adapter,
  createController,
  children,
}: {
  store: ProviderStore;
  adapter: HostAdapter | null;
  createController?: (deps: ChatControllerDeps) => ChatController;
  children: ReactNode;
}) {
  const session = useChatController({ store, adapter, createController });
  const value: ChatSessionValue = {
    ...session,
    adapter,
    store,
  };
  return (
    <ChatSessionContext.Provider value={value}>
      {children}
    </ChatSessionContext.Provider>
  );
}

export function useChatSession(): ChatSessionValue {
  const ctx = useContext(ChatSessionContext);
  if (!ctx) {
    throw new Error("useChatSession requires ChatSessionProvider");
  }
  return ctx;
}

/** Optional: panels that may render before provider is mounted. */
export function useOptionalChatSession(): ChatSessionValue | null {
  return useContext(ChatSessionContext);
}
