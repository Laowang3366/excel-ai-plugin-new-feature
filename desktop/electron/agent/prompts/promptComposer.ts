export interface PromptSection {
  key: string;
  content: string;
}

export function composePromptSections(sections: PromptSection[]): string {
  const seen = new Set<string>();
  const content: string[] = [];

  for (const section of sections) {
    if (seen.has(section.key)) continue;

    const normalized = section.content.trim();
    if (!normalized) continue;

    seen.add(section.key);
    content.push(normalized);
  }

  return content.join("\n\n");
}

export function appendPromptSections(prompt: string, sections: PromptSection[]): string {
  const suffix = composePromptSections(sections);
  if (!suffix) return prompt;

  const prefix = prompt.trim();
  return prefix ? `${prefix}\n\n${suffix}` : suffix;
}

export function renderPromptTemplate(template: string, variables: Record<string, string>): string {
  const rendered = template.replace(/\{\{([A-Z0-9_]+)\}\}/g, (_placeholder, name: string) => {
    if (!Object.prototype.hasOwnProperty.call(variables, name)) {
      throw new Error(`缺少提示词模板变量：${name}`);
    }
    return variables[name];
  });
  const unresolved = rendered.match(/\{\{[^{}]+\}\}/)?.[0];
  if (unresolved) {
    throw new Error(`存在未替换的提示词模板变量：${unresolved}`);
  }

  return rendered.trim();
}
