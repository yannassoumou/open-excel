/* global PowerPoint, console */

/* PowerPoint presentation context reading helpers */

/**
 * Read current presentation context (all slides + shape info).
 */
export async function getSlideContext() {
  try {
    const info = await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/name");
      await context.sync();

      const result = {
        slideNames: slides.items.map((s) => s.name),
        slides: [],
      };

      for (const slide of slides.items) {
        slide.load("name");
        const shapes = slide.shapes;
        shapes.load("items/name, type, placeholderType");
        await context.sync();

        const shapeInfo = shapes.items.map((shape) => ({
          name: shape.name,
          type: shape.type,
          placeholderType: shape.placeholderType,
        }));

        result.slides.push({
          name: slide.name,
          shapeCount: shapeInfo.length,
          shapes: shapeInfo,
        });
      }

      return result;
    });

    console.log(`[ppt:context] Read ${info.slides.length} slide(s):`, info.slideNames);

    const lines = [];
    for (const slide of info.slides) {
      lines.push(
        `Slide "${slide.name}": ${slide.shapeCount} shape(s) — ${slide.shapes
          .map((s) => `${s.name} (${s.type})`)
          .join(", ")}`
      );
    }

    return `Presentation has ${info.slideNames.length} slide(s): ${info.slideNames.join(
      ", "
    )}\n${lines.join("\n")}\n\nUse slide names or indices exactly as listed above.`;
  } catch (error) {
    console.warn("[ppt:context] Failed to read slide context:", error);
    return "No slide context available.";
  }
}

/**
 * Read the FULL presentation context — all slides, all shape text.
 * Used for pre-completion validation.
 */
export async function readFullPresentationContext() {
  try {
    const info = await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/name");
      await context.sync();

      const result = {
        slideNames: slides.items.map((s) => s.name),
        slides: [],
      };

      for (const slide of slides.items) {
        slide.load("name");
        await context.sync();

        const shapes = slide.shapes;
        shapes.load("items/name, type, placeholderType, text");
        await context.sync();

        const shapeInfo = shapes.items.map((shape) => ({
          name: shape.name,
          type: shape.type,
          placeholderType: shape.placeholderType,
          text: shape.text || "",
        }));

        result.slides.push({
          name: slide.name,
          shapeCount: shapeInfo.length,
          shapes: shapeInfo,
        });
      }

      return result;
    });

    const lines = [];
    lines.push(
      `Final presentation state (${info.slideNames.length} slide(s)): ${info.slideNames.join(", ")}`
    );
    for (const slide of info.slides) {
      lines.push(`  Slide "${slide.name}": ${slide.shapeCount} shape(s)`);
      for (const shape of slide.shapes) {
        const textPreview = shape.text ? shape.text.substring(0, 100).replace(/\n/g, "\\n") : "(no text)";
        lines.push(`    - "${shape.name}" (${shape.type}): ${textPreview}`);
      }
    }
    return lines.join("\n");
  } catch (error) {
    console.warn("[ppt:context] Failed to read full presentation context:", error);
    return "Could not read presentation state for validation.";
  }
}

/**
 * Lightweight presentation context reader for telemetry (shape metadata only).
 */
export async function readPresentationMetadata() {
  try {
    const info = await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/name");
      await context.sync();

      const result = [];
      for (const slide of slides.items) {
        slide.load("name");
        await context.sync();

        const shapes = slide.shapes;
        shapes.load("items/name, type");
        await context.sync();

        result.push({
          name: slide.name,
          shapeCount: shapes.items.length,
        });
      }
      return result;
    });
    return info;
  } catch (error) {
    console.warn("[ppt:context] Failed to read presentation metadata:", error);
    return null;
  }
}
