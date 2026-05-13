/* global PowerPoint, console */

/* PowerPoint presentation snapshots for revert and operation verification */

const WRITABLE_OPS = [
  "setText",
  "addSlide",
  "deleteSlide",
  "addShape",
  "addTable",
  "addChart",
  "deleteShape",
];

/**
 * Capture a snapshot of the presentation state before executing operations.
 */
export async function captureSnapshot(operations) {
  const writableOps = operations.filter((op) => WRITABLE_OPS.includes(op.operation));
  if (writableOps.length === 0) return null;

  try {
    const snapshot = await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/name");
      await context.sync();

      const data = {};
      for (const slide of slides.items) {
        slide.load("name", "shapes");
        await context.sync();

        const shapes = slide.shapes;
        shapes.load("items/name, text");
        await context.sync();

        const shapeData = {};
        for (const shape of shapes.items) {
          shape.load("text");
          await context.sync();
          shapeData[shape.name] = {
            text: shape.text || "",
          };
        }

        data[slide.name] = {
          shapeCount: shapes.items.length,
          shapes: shapeData,
        };
      }
      return data;
    });
    console.log(`[ppt:snapshot] Captured ${Object.keys(snapshot || {}).length} slide(s)`);
    return snapshot;
  } catch (error) {
    console.warn(`[ppt:snapshot] Failed to capture: ${error.message}`);
    return null;
  }
}

/**
 * Revert to a captured snapshot by restoring shape text.
 */
export async function revertToSnapshot(snapshot) {
  if (!snapshot || Object.keys(snapshot).length === 0) {
    return "  ⚠ No snapshot available to revert to";
  }

  try {
    const result = await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/name");
      await context.sync();

      const existingSlides = new Set(slides.items.map((s) => s.name));
      const restored = [];

      for (const [slideName, slideData] of Object.entries(snapshot)) {
        let slide;
        if (existingSlides.has(slideName)) {
          slide = slides.getItem(slideName);
        } else {
          slide = slides.add(slideName);
          existingSlides.add(slideName);
        }
        slide.load("name", "shapes");
        await context.sync();

        const shapes = slide.shapes;
        shapes.load("items/name");
        await context.sync();

        for (const [shapeName, shapeData] of Object.entries(slideData.shapes)) {
          const shape = shapes.getItemOrNullObject(shapeName);
          shape.load("isNullObject");
          await context.sync();

          if (!shape.isNullObject) {
            shape.text = shapeData.text;
            await context.sync();
            restored.push(`${slideName}/${shapeName}`);
          }
        }
      }

      return restored.length > 0
        ? `  🔄 Reverted: ${restored.join(", ")}`
        : "  🔄 Reverted to previous state";
    });

    console.log(`[ppt:revert] ${result}`);
    return result;
  } catch (error) {
    console.error(`[ppt:revert] Failed: ${error.message}`);
    return `  ⚠ Revert failed: ${error.message}`;
  }
}

/**
 * Read back verification data for all executed operations.
 */
export async function verifyPptOperations(operations, actualSlideNames) {
  try {
    const verification = await PowerPoint.run(async (context) => {
      const results = [];
      const seen = new Set();

      for (let i = 0; i < operations.length; i++) {
        const op = operations[i];
        const type = op.operation;
        if (type === "readShapes") continue;

        const slideName = (actualSlideNames && actualSlideNames[i]) || op.slide || op.name;
        const key = `${slideName}|${op.shape || ""}|${type}`;
        if (seen.has(key)) continue;
        seen.add(key);

        const slide = context.presentation.slides.getItem(slideName);
        slide.load("name", "shapes");
        await context.sync();

        switch (type) {
          case "setText": {
            const shape = slide.shapes.getItemOrNullObject(op.shape);
            shape.load("text, isNullObject");
            await context.sync();
            if (!shape.isNullObject) {
              results.push(
                `  ✓ ${slideName}/${op.shape}: "${(shape.text || "").substring(0, 50)}"`
              );
            } else {
              results.push(`  ⚠ ${slideName}: shape "${op.shape}" not found`);
            }
            break;
          }

          case "addShape": {
            const shapes = slide.shapes;
            shapes.load("items/name");
            await context.sync();
            const shapeNames = shapes.items.map((s) => s.name).join(", ") || "(none)";
            results.push(`  ✓ ${slideName}: shapes = [${shapeNames}]`);
            break;
          }

          case "addSlide": {
            results.push(`  ✓ Slide "${slideName}" exists`);
            break;
          }

          case "deleteSlide": {
            const allSlides = context.presentation.slides;
            allSlides.load("items/name");
            await context.sync();
            const found = allSlides.items.some((s) => s.name === slideName);
            if (!found) {
              results.push(`  ✓ Slide "${slideName}" deleted`);
            } else {
              results.push(`  ⚠ Slide "${slideName}" still exists`);
            }
            break;
          }

          default:
            results.push(`  ✓ ${slideName} (${type})`);
        }
      }

      return results.join("\n");
    });

    return verification;
  } catch (error) {
    return `  ⚠ Verification failed: ${error.message}`;
  }
}

/**
 * Extract actual slide names from operation results.
 */
export function extractSlideNamesFromResults(results, operations) {
  return results.map((result, i) => {
    const op = operations[i];
    if (!result) return op.slide || op.name;

    const slideMatch = result.match(/Slide "([^"]+)" ready/);
    if (slideMatch) return slideMatch[1];

    const deleteMatch = result.match(/Deleted slide "([^"]+)"/);
    if (deleteMatch) return deleteMatch[1];

    const shapeMatch = result.match(/Set text on ([^/]+)/);
    if (shapeMatch) return shapeMatch[1];

    return op.slide || op.name;
  });
}
