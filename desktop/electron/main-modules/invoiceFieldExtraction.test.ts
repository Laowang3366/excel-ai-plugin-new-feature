import { describe, expect, it } from "vitest";

import { buildInvoiceFieldFallback } from "./invoiceFieldExtraction";
import type { MineruParsedDocument } from "./mineruOcr";

describe("invoiceFieldExtraction", () => {
  it("extracts common invoice fields from MinerU markdown text and tables", () => {
    const documents: MineruParsedDocument[] = [{
      filename: "invoice.pdf",
      text: [
        "电子发票（增值税专用发票）",
        "发票号码：12345678",
        "开票日期：2026年07月03日",
        "购买方名称：上海示例科技有限公司",
        "购买方统一社会信用代码：91310000MA1TEST01X",
        "销售方名称：北京供应商有限公司",
        "销售方纳税人识别号：91110000MA1TEST02Y",
        "价税合计（小写） ¥106.00",
      ].join("\n"),
      rows: [
        ["项目", "金额", "税额"],
        ["服务费", "100.00", "6.00"],
        ["合计金额", "100.00"],
        ["合计税额", "6.00"],
      ],
    }];

    const result = buildInvoiceFieldFallback(documents);

    expect(result.fields).toMatchObject({
      发票号码: "12345678",
      开票日期: "2026年07月03日",
      购买方名称: "上海示例科技有限公司",
      购买方税号: "91310000MA1TEST01X",
      销售方名称: "北京供应商有限公司",
      销售方税号: "91110000MA1TEST02Y",
      金额: "100.00",
      税额: "6.00",
      价税合计: "¥106.00",
    });
    expect(result.rows[0]).toContain("发票号码");
    expect(result.invoices[0].fields["发票号码"]).toBe("12345678");
  });
});
