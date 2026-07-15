import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import { parseAndValidateToolArguments } from "./toolSchema";

function parameters(name: string) {
  const definition = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === name);
  if (!definition) throw new Error(`missing tool definition: ${name}`);
  return definition.parameters;
}

function apply(operation: string, params: Record<string, unknown>, outputPath?: string) {
  return {
    app: "presentation",
    action: "edit",
    operation,
    filePath: "C:/presentations/report.pptx",
    target: "slide:2",
    ...(outputPath ? { outputPath } : {}),
    params,
  };
}

describe("PowerPoint playback and inspection parameter schemas", () => {
  it("accepts strict presentation inspections and validation conditions", () => {
    const cases = [
      ["inspectPresentationTheme", { host: "powerpoint" }],
      ["inspectSlideElements", { allSlides: true, countPath: "summary.shapeCount", minCount: 1 }],
      ["inspectAnimations", { allSlides: false }],
      ["inspectSpeakerNotes", { allSlides: true, containsText: "Summary" }],
    ] as const;

    for (const [operation, params] of cases) {
      const args = {
        app: "presentation",
        operation,
        filePath: "C:/presentations/report.pptx",
        params,
      };
      for (const tool of ["office.action.inspect", "office.action.validate"]) {
        expect(
          parseAndValidateToolArguments(JSON.stringify(args), parameters(tool)).error,
        ).toBeUndefined();
      }
    }
  });

  it("accepts explicit animation rules and workflow steps", () => {
    const args = apply("configureAnimations", {
      host: "powerpoint",
      clearExisting: true,
      effects: [
        {
          category: "entrance",
          effect: "fade",
          shapeName: "Title 1",
          trigger: "onClick",
          duration: 0.5,
        },
        {
          category: "path",
          effect: "fly",
          shapeNames: ["Chart 2", "TextBox 3"],
          trigger: "afterPrevious",
          pathX: 0.2,
          pathY: 0,
        },
      ],
    });

    expect(
      parseAndValidateToolArguments(JSON.stringify(args), parameters("office.action.apply")).error,
    ).toBeUndefined();
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ steps: [args] }),
        parameters("office.workflow.run"),
      ).error,
    ).toBeUndefined();
  });

  it("rejects implicit all-shape animation and malformed nested rules", () => {
    const schema = parameters("office.action.apply");
    const invalid = [
      {},
      { clearExisting: true },
      { effects: [] },
      { effects: [{ category: "entrance", effect: "fade" }] },
      { effects: [{ category: "entrance", effect: "bounce", shapeName: "Title 1" }] },
      { effects: [{ category: "entrance", effect: "pulse", shapeName: "Title 1" }] },
      {
        effects: [
          {
            category: "entrance",
            effect: "fade",
            shapeName: "Title 1",
            shapeNames: ["Title 1"],
          },
        ],
      },
      {
        effects: [{ category: "entrance", effect: "fade", shapeName: "Title 1", easing: "linear" }],
      },
    ];

    for (const params of invalid) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("configureAnimations", params)), schema)
          .error,
      ).toBeDefined();
    }
  });

  it("requires an explicit slide-show type and real transition fields", () => {
    const schema = parameters("office.action.apply");
    const valid = apply("configureSlideShow", {
      showType: "kiosk",
      autoPlay: true,
      loop: true,
      transition: "wipe",
      advanceAfter: 8,
      transitionDuration: 1,
    });
    expect(parseAndValidateToolArguments(JSON.stringify(valid), schema).error).toBeUndefined();

    for (const params of [{}, { transition: "fade" }, { showType: "fullScreen" }]) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("configureSlideShow", params)), schema)
          .error,
      ).toBeDefined();
    }
  });

  it("accepts single and batch speaker notes but rejects ambiguous or incomplete payloads", () => {
    const schema = parameters("office.action.apply");
    const valid = [
      { text: "Explain the revenue bridge.", append: true },
      {
        notesBySlide: [
          { slideIndex: 1, text: "Opening" },
          { slideIndex: 2, text: "Results", append: true },
        ],
      },
    ];
    for (const params of valid) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("setSpeakerNotes", params)), schema)
          .error,
      ).toBeUndefined();
    }

    const invalid = [
      {},
      { notesBySlide: [] },
      { notesBySlide: [{ slideIndex: 0, text: "Invalid" }] },
      { notesBySlide: [{ slideIndex: 1 }] },
      { text: "One", notesBySlide: [{ slideIndex: 1, text: "Two" }] },
    ];
    for (const params of invalid) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("setSpeakerNotes", params)), schema)
          .error,
      ).toBeDefined();
    }
  });

  it("requires an explicit handout layout and keeps outputPath at the action level", () => {
    const schema = parameters("office.action.apply");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify(
          apply(
            "exportHandouts",
            { host: "powerpoint", layout: "six", includeNotes: false },
            "C:/presentations/report-handouts.pdf",
          ),
        ),
        schema,
      ).error,
    ).toBeUndefined();

    for (const params of [
      {},
      { layout: "five" },
      { layout: "three", outputPath: "C:/presentations/wrong.pdf" },
    ]) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("exportHandouts", params)), schema)
          .error,
      ).toBeDefined();
    }
  });
});
