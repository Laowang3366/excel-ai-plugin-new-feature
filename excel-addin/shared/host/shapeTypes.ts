/** Geometric shape whitelist (Phase15 MVP; not full GeometricShapeType enum). */
export type GeometricShapeType =
  | "rectangle"
  | "ellipse"
  | "triangle"
  | "diamond"
  | "rightArrow";

export const GEOMETRIC_SHAPE_TYPES: readonly GeometricShapeType[] = [
  "rectangle",
  "ellipse",
  "triangle",
  "diamond",
  "rightArrow",
];

export function isGeometricShapeType(value: unknown): value is GeometricShapeType {
  return (
    typeof value === "string" &&
    (GEOMETRIC_SHAPE_TYPES as readonly string[]).includes(value)
  );
}

export type ShapeCreateKind = "geometric" | "textBox";

export interface ShapeInfo {
  name: string;
  sheetName: string;
  /** Host Shape.type (e.g. GeometricShape, Image, Group, Line). */
  type: string;
  geometricShapeType?: string | null;
  left: number;
  top: number;
  width: number;
  height: number;
  visible?: boolean;
  /** null when textFrame.hasText is false. */
  text?: string | null;
}

export type ShapeCreateInput =
  | {
      sheetName: string;
      kind: "geometric";
      geometricType: GeometricShapeType;
      name?: string;
      left?: number;
      top?: number;
      width?: number;
      height?: number;
    }
  | {
      sheetName: string;
      kind: "textBox";
      text?: string;
      name?: string;
      left?: number;
      top?: number;
      width?: number;
      height?: number;
    };

export interface ShapeUpdateInput {
  sheetName: string;
  shapeName: string;
  newName?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  text?: string;
  visible?: boolean;
}
