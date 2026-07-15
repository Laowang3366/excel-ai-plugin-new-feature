import { access, mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  DotNetOfficeActionBridge as OfficeComActionBridge,
  applyPresentationAdvancedAction,
  disposeOfficeWorker,
} from "./officeWorkerSmokeHelpers";
import type { OfficeActionInput } from "../electron/agent/tools/officeCore/types";

async function main(): Promise<void> {
  const tempDir = await mkdtemp(path.join(os.tmpdir(), "wengge-presentation-smoke-"));
  const presentationPath = path.join(tempDir, "advanced.pptx");
  const templatePath = path.join(tempDir, "brand-template.pptx");
  const handoutPath = path.join(tempDir, "notes.pdf");
  const logoPath = path.join(tempDir, "logo.png");
  const keepArtifacts = process.env.KEEP_PRESENTATION_SMOKE === "1";
  const partialFailureOnly = process.env.PRESENTATION_SMOKE_PARTIAL_FAILURE_ONLY === "1";
  const requestedHost = (process.env.PRESENTATION_SMOKE_HOST || "").trim().toLowerCase();
  const operationFilter = new Set(
    (process.env.PRESENTATION_SMOKE_OPERATIONS || "")
      .split(",")
      .map((operation) => operation.trim())
      .filter(Boolean),
  );

  try {
    await createFixture(presentationPath, templatePath, logoPath);
    const bridge = new OfficeComActionBridge();
    const actions: OfficeActionInput[] = [
      {
        app: "presentation",
        action: "inspect",
        operation: "inspectPresentationTheme",
        filePath: presentationPath,
      },
      {
        app: "presentation",
        action: "style",
        operation: "applyMasterBranding",
        filePath: presentationPath,
        params: {
          fontName: "Microsoft YaHei",
          templatePath,
          accentColor: "2563EB",
          backgroundColor: "FFFFFF",
          logoPath,
          logoWidth: 24,
          footerText: "Office AI 演示",
          showSlideNumber: true,
          themeColors: [
            { index: 5, value: "2563EB" },
            { index: 6, value: "16A34A" },
          ],
        },
      },
      {
        app: "presentation",
        action: "inspect",
        operation: "inspectPresentationTheme",
        filePath: presentationPath,
        params: {
          expectedThemeColors: [
            { index: 5, value: "2563EB" },
            { index: 6, value: "16A34A" },
          ],
        },
      },
      {
        app: "presentation",
        action: "insert",
        operation: "insertChart",
        filePath: presentationPath,
        target: "slide:2",
        params: { name: "Revenue Chart", chartType: "column" },
      },
      {
        app: "presentation",
        action: "insert",
        operation: "insertTable",
        filePath: presentationPath,
        target: "slide:2",
        params: {
          name: "KPI Table",
          values: [
            ["指标", "结果"],
            ["收入增长", "18%"],
          ],
          left: 80,
          top: 360,
          width: 520,
          height: 120,
        },
      },
      {
        app: "presentation",
        action: "insert",
        operation: "replacePictureSlot",
        filePath: presentationPath,
        target: "slide:2",
        params: {
          name: "Product Image",
          imagePath: logoPath,
          left: 640,
          top: 140,
          width: 120,
          height: 120,
          preserveAspectRatio: true,
        },
      },
      {
        app: "presentation",
        action: "inspect",
        operation: "inspectSlideElements",
        filePath: presentationPath,
        params: { allSlides: true },
      },
      {
        app: "presentation",
        action: "style",
        operation: "layoutElements",
        filePath: presentationPath,
        params: {
          allSlides: true,
          mode: "grid",
          columns: 2,
          margin: 42,
          gap: 18,
          rowHeight: 140,
          resize: true,
          fitToSlide: true,
          edits: [
            {
              shapeName: "KPI Table",
              tableCells: [
                { row: 2, column: 2, text: "20%", fontName: "Microsoft YaHei", fontSize: 16 },
              ],
            },
            { shapeName: "Revenue Chart", chart: { title: "季度收入", hasLegend: false } },
            {
              shapeName: "Product Image",
              preserveAspectRatio: true,
              crop: { left: 0, right: 0, top: 0, bottom: 0 },
            },
          ],
        },
      },
      {
        app: "presentation",
        action: "inspect",
        operation: "inspectSlideElements",
        filePath: presentationPath,
        params: { allSlides: true },
      },
      {
        app: "presentation",
        action: "edit",
        operation: "configureAnimations",
        filePath: presentationPath,
        target: "slide:1",
        params: {
          clearExisting: true,
          effects: [
            { category: "entrance", effect: "fade", trigger: "onClick", duration: 0.4 },
            { category: "emphasis", effect: "growShrink", trigger: "withPrevious", duration: 0.3 },
            { category: "exit", effect: "wipe", trigger: "afterPrevious", duration: 0.4 },
            {
              category: "path",
              effect: "appear",
              trigger: "afterPrevious",
              duration: 0.5,
              pathX: 0.08,
              pathY: 0,
            },
          ],
        },
      },
      {
        app: "presentation",
        action: "inspect",
        operation: "inspectAnimations",
        filePath: presentationPath,
        target: "slide:1",
      },
      {
        app: "presentation",
        action: "edit",
        operation: "configureSlideShow",
        filePath: presentationPath,
        params: {
          allSlides: true,
          showType: "speaker",
          autoPlay: true,
          loop: true,
          transition: "fade",
          advanceAfter: 3,
          transitionDuration: 0.5,
        },
      },
      {
        app: "presentation",
        action: "edit",
        operation: "setSpeakerNotes",
        filePath: presentationPath,
        params: {
          notesBySlide: [
            { slideIndex: 1, text: "介绍季度经营汇报，说明收入增长和重点项目。" },
            { slideIndex: 2, text: "讲解业绩概览，突出收入增长和交付质量。" },
          ],
        },
      },
      {
        app: "presentation",
        action: "inspect",
        operation: "inspectSpeakerNotes",
        filePath: presentationPath,
        params: { allSlides: true },
      },
      {
        app: "presentation",
        action: "snapshot",
        operation: "exportHandouts",
        filePath: presentationPath,
        outputPath: handoutPath,
        params: { includeNotes: true },
      },
    ];
    const hostActions = requestedHost
      ? actions.map((action) => ({ ...action, params: { ...action.params, host: requestedHost } }))
      : actions;
    const selectedActions = partialFailureOnly
      ? []
      : operationFilter.size > 0
        ? hostActions.filter((action) => operationFilter.has(action.operation))
        : hostActions;
    const results = [];

    for (const [index, action] of selectedActions.entries()) {
      process.stdout.write(`Testing ${action.operation}\n`);
      if (keepArtifacts)
        await writeFile(
          path.join(tempDir, `${index}-${action.operation}.json`),
          JSON.stringify(action, null, 2),
          "utf8",
        );
      const result = await bridge.executeAction(action);
      results.push({ operation: action.operation, status: result.status, error: result.error });
      if (result.status !== "done")
        throw new Error(`${action.operation}: ${result.error || result.summary}`);
      verifyResult(action, result.data);
    }
    if (
      (partialFailureOnly ||
        selectedActions.some((action) => action.operation === "layoutElements")) &&
      requestedHost !== "wps"
    ) {
      await verifyMissingShapeFails(bridge, presentationPath, requestedHost);
      results.push({
        operation: "layoutElements:missing-shape",
        status: "partial_failure",
        error: undefined,
      });
    }
    if (selectedActions.some((action) => action.operation === "exportHandouts"))
      await access(handoutPath);
    process.stdout.write(`${JSON.stringify({ ok: true, results, handoutPath }, null, 2)}\n`);
  } finally {
    await disposeOfficeWorker();
    if (keepArtifacts) process.stdout.write(`Presentation smoke artifacts: ${tempDir}\n`);
    else await rm(tempDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  }
}

async function verifyMissingShapeFails(
  bridge: DotNetOfficeActionBridge,
  presentationPath: string,
  host: string,
): Promise<void> {
  try {
    await bridge.executeAction({
      app: "presentation",
      action: "style",
      operation: "layoutElements",
      filePath: presentationPath,
      params: {
        ...(host ? { host } : {}),
        mode: "none",
        edits: [{ shapeName: "__WENGGE_MISSING_SHAPE__", left: 10 }],
      },
    });
  } catch (error) {
    if ((error as { code?: string }).code === "partial_failure") return;
    throw error;
  }
  throw new Error("layoutElements 找不到形状时仍错误返回成功");
}

function verifyResult(action: OfficeActionInput, data: unknown): void {
  const operationData = asRecord(data);
  const requestedHost = String(action.params?.host || "").toLowerCase();
  const progId = String(operationData.progId || "").toLowerCase();
  if (requestedHost === "wps" && !progId.includes("wpp"))
    throw new Error(`未使用 WPS 演示 COM: ${operationData.progId || "unknown"}`);
  if (requestedHost === "powerpoint" && progId !== "powerpoint.application")
    throw new Error(`未使用 PowerPoint COM: ${operationData.progId || "unknown"}`);
  if (action.operation === "inspectPresentationTheme") {
    const theme = asRecord(operationData.theme);
    if (Number(theme.slideCount) < 2 || !Array.isArray(theme.designs) || theme.designs.length < 1)
      throw new Error("主题母版检查失败");
    const expectedColors = Array.isArray(action.params?.expectedThemeColors)
      ? action.params.expectedThemeColors.map(asRecord)
      : [];
    const colors = (theme.designs as unknown[])
      .flatMap((design) =>
        Array.isArray(asRecord(design).colors) ? (asRecord(design).colors as unknown[]) : [],
      )
      .map(asRecord);
    for (const expected of expectedColors) {
      const expectedRgb = colorHexToOle(String(expected.value || ""));
      if (
        !colors.some(
          (color) =>
            Number(color.index) === Number(expected.index) && Number(color.rgb) === expectedRgb,
        )
      ) {
        throw new Error(`主题色回读失败: ${JSON.stringify(colors)}`);
      }
    }
  }
  if (
    action.operation === "applyMasterBranding" &&
    Number(asRecord(operationData.updated).masters) < 1
  )
    throw new Error("母版品牌更新失败");
  if (action.operation === "applyMasterBranding") {
    if (operationData.appliedTemplate !== action.params?.templatePath)
      throw new Error("旧 PPT 模板转换失败");
    const designs = Array.isArray(asRecord(operationData.theme).designs)
      ? (asRecord(operationData.theme).designs as unknown[])
      : [];
    const colors = designs
      .flatMap((design) =>
        Array.isArray(asRecord(design).colors) ? (asRecord(design).colors as unknown[]) : [],
      )
      .map(asRecord);
    if (requestedHost === "wps") {
      if (operationData.themePackageFallback !== true)
        throw new Error("WPS 主题色未触发 Open XML 兼容更新");
    } else if (
      !colors.some((color) => Number(color.index) === 5 && Number(color.rgb) === 15426341)
    ) {
      throw new Error(`主题色更新失败: ${JSON.stringify(colors)}`);
    }
  }
  if (action.operation === "inspectSlideElements") {
    const summary = asRecord(operationData.summary);
    if (Number(summary.slideCount) < 2 || Number(summary.shapeCount) < 2)
      throw new Error(`元素检查失败: ${JSON.stringify(summary)}`);
    const slides = Array.isArray(operationData.slides) ? operationData.slides.map(asRecord) : [];
    const shapes = slides
      .flatMap((slide) => (Array.isArray(slide.shapes) ? (slide.shapes as unknown[]) : []))
      .map(asRecord);
    for (const typeName of ["chart", "table", "picture"]) {
      if (!shapes.some((shape) => shape.typeName === typeName))
        throw new Error(`元素检查缺少 ${typeName}`);
    }
  }
  if (action.operation === "configureAnimations") {
    const snapshot = asRecord(operationData.snapshot);
    if (!Array.isArray(snapshot.effects) || snapshot.effects.length < 4)
      throw new Error("四类动画配置失败");
    const animated = Array.isArray(operationData.animated)
      ? operationData.animated.map(asRecord)
      : [];
    for (const category of ["entrance", "emphasis", "exit", "path"]) {
      if (!animated.some((item) => item.category === category))
        throw new Error(`动画配置缺少 ${category}`);
    }
    const effects = snapshot.effects.map(asRecord);
    if (!effects.some((effect) => effect.exit === true)) throw new Error("退出动画未生效");
    if (
      !effects.some(
        (effect) =>
          Array.isArray(effect.behaviors) &&
          effect.behaviors
            .map(asRecord)
            .some((behavior) => Number(behavior.type) === 1 && Math.abs(Number(behavior.byX)) > 0),
      )
    )
      throw new Error("路径动画行为未生效");
  }
  if (action.operation === "inspectAnimations") {
    const animations = Array.isArray(operationData.animations)
      ? operationData.animations.map(asRecord)
      : [];
    if (
      animations.length < 1 ||
      !Array.isArray(animations[0].effects) ||
      animations[0].effects.length < 4
    )
      throw new Error("动画检查失败");
  }
  if (action.operation === "configureSlideShow" && asRecord(operationData.slideShow).loop !== true)
    throw new Error("循环放映配置失败");
  if (
    action.operation === "setSpeakerNotes" &&
    (!Array.isArray(operationData.updatedSlides) || operationData.updatedSlides.length !== 2)
  )
    throw new Error("批量备注写入失败");
  if (
    action.operation === "inspectSpeakerNotes" &&
    Number(asRecord(operationData.summary).missingNotes) !== 0
  )
    throw new Error("备注检查发现缺失讲稿");
}

function asRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : {};
}

function colorHexToOle(value: string): number {
  const hex = value.replace(/^#/, "");
  if (!/^[0-9a-f]{6}$/i.test(hex)) return Number.NaN;
  return (
    Number.parseInt(hex.slice(0, 2), 16) +
    Number.parseInt(hex.slice(2, 4), 16) * 256 +
    Number.parseInt(hex.slice(4, 6), 16) * 65536
  );
}

async function createFixture(
  presentationPath: string,
  templatePath: string,
  logoPath: string,
): Promise<void> {
  await writeFile(
    logoPath,
    Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
      "base64",
    ),
  );
  const created = await applyPresentationAdvancedAction({
    operation: "createPresentation",
    filePath: presentationPath,
    params: { title: "季度经营汇报", subtitle: "收入增长与重点项目" },
  });
  if (created.status !== "done") throw new Error(created.error || created.summary);
  const template = await applyPresentationAdvancedAction({
    operation: "createPresentation",
    filePath: templatePath,
    params: { title: "品牌模板", subtitle: "统一版式" },
  });
  if (template.status !== "done") throw new Error(template.error || template.summary);
  const added = await applyPresentationAdvancedAction({
    operation: "addSlides",
    filePath: presentationPath,
    params: { slides: [{ title: "业绩概览", body: "收入增长 18%\n重点项目按期交付" }] },
  });
  if (added.status !== "done") throw new Error(added.error || added.summary);
}

void main().catch((error) => {
  process.stderr.write(
    `${error instanceof Error ? error.stack || error.message : String(error)}\n`,
  );
  process.exitCode = 1;
});
