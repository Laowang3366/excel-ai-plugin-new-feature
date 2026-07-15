export interface JsonResourceBudget {
  maxDepth: number;
  maxNodes: number;
  maxStringChars: number;
  maxArrayItems: number;
  maxObjectProperties: number;
  maxSerializedBytes: number;
}

export interface JsonResourceBudgetViolation {
  path: Array<string | number>;
  message: string;
}

export const DEFAULT_IPC_JSON_RESOURCE_BUDGET: JsonResourceBudget = {
  maxDepth: 16,
  maxNodes: 100_000,
  maxStringChars: 1_048_576,
  maxArrayItems: 20_000,
  maxObjectProperties: 512,
  maxSerializedBytes: 4 * 1024 * 1024,
};

interface PendingValue {
  value: unknown;
  path: Array<string | number>;
  depth: number;
}

function jsonEscapedUtf8ByteLength(value: string): number {
  let bytes = 0;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code === 0x22 || code === 0x5c) bytes += 2;
    else if (code <= 0x1f) bytes += 6;
    else if (code < 0x80) bytes += 1;
    else if (code < 0x800) bytes += 2;
    else if (code >= 0xd800 && code <= 0xdbff && index + 1 < value.length) {
      const next = value.charCodeAt(index + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        bytes += 4;
        index += 1;
      } else {
        bytes += 3;
      }
    } else bytes += 3;
  }
  return bytes;
}

function isPlainObject(value: object): value is Record<string, unknown> {
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

export function inspectJsonResourceBudget(
  value: unknown,
  budget: JsonResourceBudget = DEFAULT_IPC_JSON_RESOURCE_BUDGET,
): JsonResourceBudgetViolation | null {
  const pending: PendingValue[] = [{ value, path: [], depth: 0 }];
  const seen = new WeakSet<object>();
  let nodes = 0;
  let serializedBytes = 0;

  while (pending.length > 0) {
    const current = pending.pop()!;
    nodes += 1;
    if (nodes > budget.maxNodes) {
      return { path: current.path, message: `JSON 节点数不能超过 ${budget.maxNodes}` };
    }
    if (current.depth > budget.maxDepth) {
      return { path: current.path, message: `JSON 嵌套深度不能超过 ${budget.maxDepth}` };
    }

    const currentValue = current.value;
    if (currentValue === null) {
      serializedBytes += 4;
    } else if (typeof currentValue === "string") {
      if (currentValue.length > budget.maxStringChars) {
        return {
          path: current.path,
          message: `JSON 字符串不能超过 ${budget.maxStringChars} 个字符`,
        };
      }
      serializedBytes += jsonEscapedUtf8ByteLength(currentValue) + 2;
    } else if (typeof currentValue === "number") {
      if (!Number.isFinite(currentValue)) {
        return { path: current.path, message: "JSON 数字必须是有限值" };
      }
      serializedBytes += String(currentValue).length;
    } else if (typeof currentValue === "boolean") {
      serializedBytes += currentValue ? 4 : 5;
    } else if (typeof currentValue === "object") {
      if (seen.has(currentValue)) {
        return { path: current.path, message: "JSON 不能包含循环引用" };
      }
      seen.add(currentValue);

      if (Array.isArray(currentValue)) {
        if (currentValue.length > budget.maxArrayItems) {
          return {
            path: current.path,
            message: `JSON 数组不能超过 ${budget.maxArrayItems} 项`,
          };
        }
        serializedBytes += currentValue.length > 0 ? currentValue.length + 1 : 2;
        for (let index = currentValue.length - 1; index >= 0; index -= 1) {
          pending.push({
            value: currentValue[index],
            path: [...current.path, index],
            depth: current.depth + 1,
          });
        }
      } else {
        if (!isPlainObject(currentValue)) {
          return { path: current.path, message: "JSON 仅支持普通对象" };
        }
        const entries = Object.entries(currentValue);
        if (entries.length > budget.maxObjectProperties) {
          return {
            path: current.path,
            message: `JSON 对象字段不能超过 ${budget.maxObjectProperties} 个`,
          };
        }
        serializedBytes += entries.length > 0 ? entries.length + 1 : 2;
        for (let index = entries.length - 1; index >= 0; index -= 1) {
          const [key, childValue] = entries[index];
          if (key.length > budget.maxStringChars) {
            return {
              path: [...current.path, key],
              message: `JSON 字段名不能超过 ${budget.maxStringChars} 个字符`,
            };
          }
          serializedBytes += jsonEscapedUtf8ByteLength(key) + 3;
          pending.push({
            value: childValue,
            path: [...current.path, key],
            depth: current.depth + 1,
          });
        }
      }
    } else {
      return { path: current.path, message: "值必须是可序列化的 JSON" };
    }

    if (serializedBytes > budget.maxSerializedBytes) {
      return {
        path: current.path,
        message: `JSON 序列化大小不能超过 ${budget.maxSerializedBytes} 字节`,
      };
    }
  }

  return null;
}
