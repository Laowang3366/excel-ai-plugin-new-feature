import { withExcel } from "./officeJsRuntime";
import type { ExcelShape } from "./officeJsShapeFacade";
import type {
  GeometricShapeType,
  ShapeCreateInput,
  ShapeInfo,
  ShapeUpdateInput,
} from "./shapeTypes";
import type { HostResult } from "./types";

const GEOMETRIC_OFFICE: Record<GeometricShapeType, string> = {
  rectangle: "Rectangle",
  ellipse: "Ellipse",
  triangle: "Triangle",
  diamond: "Diamond",
  rightArrow: "RightArrow",
};

function mapTypeLabel(raw: string | undefined | null): string {
  return String(raw ?? "Unsupported");
}

async function readShapeInfo(
  context: { sync(): Promise<void> },
  sheetName: string,
  shape: ExcelShape,
): Promise<ShapeInfo> {
  shape.load("name,type,geometricShapeType,left,top,width,height,visible");
  shape.textFrame.load("hasText");
  await context.sync();

  let text: string | null = null;
  if (shape.textFrame.hasText === true) {
    shape.textFrame.textRange.load("text");
    await context.sync();
    text = shape.textFrame.textRange.text ?? "";
  }

  return {
    name: shape.name,
    sheetName,
    type: mapTypeLabel(shape.type),
    geometricShapeType:
      shape.geometricShapeType == null ? null : String(shape.geometricShapeType),
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
    visible: shape.visible,
    text,
  };
}

function applyGeometry(
  shape: ExcelShape,
  input: {
    name?: string;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
  },
): void {
  if (input.name != null) shape.name = input.name;
  if (input.left != null) shape.left = input.left;
  if (input.top != null) shape.top = input.top;
  if (input.width != null) shape.width = input.width;
  if (input.height != null) shape.height = input.height;
}

export async function officeJsListShapes(
  sheetName?: string,
): Promise<HostResult<ShapeInfo[]>> {
  return withExcel("shape.list", async (context) => {
    const result: ShapeInfo[] = [];
    if (sheetName) {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      sheet.load("name");
      sheet.shapes.load("items");
      await context.sync();
      for (const shape of sheet.shapes.items) {
        result.push(await readShapeInfo(context, sheet.name, shape));
      }
      return result;
    }
    context.workbook.worksheets.load("items/name");
    await context.sync();
    for (const sheet of context.workbook.worksheets.items) {
      sheet.shapes.load("items");
      await context.sync();
      for (const shape of sheet.shapes.items) {
        result.push(await readShapeInfo(context, sheet.name, shape));
      }
    }
    return result;
  });
}

export async function officeJsCreateShape(
  input: ShapeCreateInput,
): Promise<HostResult<ShapeInfo>> {
  return withExcel("shape.create", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    let shape: ExcelShape;
    if (input.kind === "geometric") {
      shape = sheet.shapes.addGeometricShape(GEOMETRIC_OFFICE[input.geometricType]);
    } else if (input.text === undefined) {
      // Official optional text: omit argument when not provided.
      shape = sheet.shapes.addTextBox();
    } else {
      shape = sheet.shapes.addTextBox(input.text);
    }
    applyGeometry(shape, input);
    await context.sync();
    return readShapeInfo(context, input.sheetName, shape);
  });
}

export async function officeJsDeleteShape(
  sheetName: string,
  shapeName: string,
): Promise<HostResult<{ deleted: string }>> {
  return withExcel("shape.delete", async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.shapes.getItem(shapeName).delete();
    await context.sync();
    // Read back host collection; do not echo input without observation.
    sheet.shapes.load("items/name");
    await context.sync();
    const stillPresent = sheet.shapes.items.some((item) => item.name === shapeName);
    if (stillPresent) {
      throw new Error(`shape still present after delete: ${shapeName}`);
    }
    return { deleted: shapeName };
  });
}

export async function officeJsUpdateShape(
  input: ShapeUpdateInput,
): Promise<HostResult<ShapeInfo>> {
  return withExcel("shape.update", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const shape = sheet.shapes.getItem(input.shapeName);
    if (input.newName != null) shape.name = input.newName;
    if (input.left != null) shape.left = input.left;
    if (input.top != null) shape.top = input.top;
    if (input.width != null) shape.width = input.width;
    if (input.height != null) shape.height = input.height;
    if (input.visible != null) shape.visible = input.visible;
    if (input.text != null) {
      shape.textFrame.textRange.text = input.text;
    }
    await context.sync();
    const name = input.newName ?? input.shapeName;
    const updated = sheet.shapes.getItem(name);
    return readShapeInfo(context, input.sheetName, updated);
  });
}
