import { describe, expect, test } from "vitest";
import { appendFolderContext, buildSystemPrompt } from "./systemPrompt";

describe("buildSystemPrompt", () => {
  test("tells formula assistant how to handle formula length and text results", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain("8192");
    expect(prompt).toContain("公式长度");
    expect(prompt).toContain("未超过");
    expect(prompt).toContain("文本");
    expect(prompt).toContain("公式使用错误");
  });

  test("guides final replies toward rendered tables without raw bold markers", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain("Markdown 表格");
    expect(prompt).toContain("避免输出原始表格分隔线文本");
    expect(prompt).toContain("不要用 `**`");
  });

  test("keeps Office document tool guidance", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain("Word 与 PowerPoint 文档");
    expect(prompt).toContain("word.open");
    expect(prompt).toContain("presentation.open");
    expect(prompt).toContain("office.script.execute");
    expect(prompt).toContain("python.execute");
    expect(prompt).toContain("不要用 `shell.execute` 拼 `python -c`");
  });

  test("guides Office visual optimization through Open XML first", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain("Open XML 优先");
    expect(prompt).toContain("office.action.inspect");
    expect(prompt).toContain("office.action.apply");
    expect(prompt).not.toContain("office.visual.snapshot");
    expect(prompt).not.toContain("office.table.applyStyle");
    expect(prompt).toContain("COM 自动兜底");
    expect(prompt).toContain("preferEngine:\"com\"");
  });

  test("guides advanced Office work through unified actions", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain("office.action.apply");
    expect(prompt).toContain("done / unsupported / needsCom / failed");
    expect(prompt).toContain("Excel/Word/PPT 高级操作优先使用统一 Office action");
    expect(prompt).toContain("打开/open 不等于文件级编辑");
    expect(prompt).toContain("不要因为 word.open 或 presentation.open 的 COM 打开失败就改用 python-pptx、pip 或 shell 拼命令");
  });

  test("maps tool responsibilities and overlap boundaries for model routing", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain("工具职责与优先级速查表");
    expect(prompt).toContain("重叠工具边界");
    expect(prompt).toContain("office.action.apply` vs `range.write");
    expect(prompt).toContain("已连接 Excel 且操作当前工作簿时优先 `range.write`");
    expect(prompt).toContain("文件级创建优先 `office.action.apply`");
    expect(prompt).toContain("`office.script.execute` 只管 Word/PPT COM");
    expect(prompt).toContain("多行 Python/文件处理用 `python.execute`");
    expect(prompt).toContain("知识库查项目/文件知识；memory 查用户偏好和长期记忆");
    expect(prompt).toContain("memory.delete");
    expect(prompt).toContain("先用 list/search 确认 memoryId");
    expect(prompt).toContain("只停用长期记忆，不删除知识库文件或知识来源");
  });

  test("does not duplicate the Office data read/write routing section", () => {
    const prompt = buildSystemPrompt();

    expect(prompt.match(/### 数据读写（注意区分读和写！）/g)).toHaveLength(1);
  });

  test("keeps attachment routing aligned with Office connection checks", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain('Excel 文件（.xlsx/.xls/.csv）：先 office.connection.status({ app:"excel" })');
    expect(prompt).toContain('Word 文件（.doc/.docx）：先 office.connection.status({ app:"word" })');
    expect(prompt).toContain('PowerPoint 文件（.ppt/.pptx）：先 office.connection.status({ app:"presentation" })');
    expect(prompt).toContain("文件级读取/编辑优先 office.action.*");
    expect(prompt).not.toContain("Excel 文件（.xlsx/.xls/.csv）：用 workbook.open 打开后操作");
  });

  test("requires Office connection preflight for modules, follow-ups and visual design tasks", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain("Office 连接预检铁律");
    expect(prompt).toContain("普通对话、功能模块、用户追问、附件任务");
    expect(prompt).toContain("视觉设计");
    expect(prompt).toContain("第一步必须调用 `office.connection.status`");
    expect(prompt).toContain("不要无故在桌面、下载或项目目录生成大量新文件");
    expect(prompt).toContain("优先检查并修改当前打开文件");
  });

  test("routes invoice recognition module through OCR tool and Excel write", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain("【功能模块：发票识别】");
    expect(prompt).toContain('ocr.parseDocument');
    expect(prompt).toContain('mode:"invoice"');
    expect(prompt).toContain("发票号码");
    expect(prompt).toContain("range.write");
    expect(prompt).toContain("写入后回读验证一次");
  });

  test("guides formula assistant to validate dynamic array spill ranges", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain('expand:"spill"');
    expect(prompt).toContain("动态数组公式必须用");
    expect(prompt).toContain("允许输出测试报告");
  });

  test("keeps shell approval and sandbox guidance", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain("shell.execute 受命令安全策略");
    expect(prompt).toContain("prompt");
    expect(prompt).toContain("forbidden");
  });

  test("appends folder context with Office open guidance", () => {
    const prompt = appendFolderContext("base", "D:\\work", "work", [
      { fileName: "demo.xlsx", filePath: "D:\\work\\demo.xlsx", size: 2048 },
    ]);

    expect(prompt).toContain("当前工作文件夹");
    expect(prompt).toContain("demo.xlsx (2KB)");
    expect(prompt).toContain("文件级读取/编辑优先使用 office.action.*");
  });
});
