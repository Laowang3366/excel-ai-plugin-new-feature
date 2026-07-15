import type {
  HostEnvironment,
  PreferredLanguage,
  ReferenceSampleMode,
} from "../../utils/taskComposerPayloads";

export interface CodeTaskDraft {
  dataSourceRanges: string[];
  dataSourceInput: string;
  referenceSampleRange: string;
  referenceSampleMode: ReferenceSampleMode;
  outputRange: string;
  hostEnvironment: HostEnvironment;
  preferredLanguage: PreferredLanguage;
  task: string;
}

export const CODE_TASK_HOST_OPTIONS: Array<{
  value: HostEnvironment;
  label: string;
}> = [
  { value: "wps", label: "WPS" },
  { value: "microsoft_excel", label: "Microsoft Excel" },
];

export const CODE_TASK_LANGUAGE_OPTIONS: Array<{
  value: PreferredLanguage;
  label: string;
}> = [
  { value: "auto", label: "自动" },
  { value: "js", label: "JS" },
  { value: "vba", label: "VBA" },
  { value: "python", label: "Python" },
];
