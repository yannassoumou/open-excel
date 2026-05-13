/* global PowerPoint, console */

/* PowerPoint operation registry and execution engine */

export const PPT_OPERATION_REGISTRY = {
  addSlide: {
    required: ["name"],
    validate: (op) => typeof op.name === "string" && op.name.length > 0,
  },
  deleteSlide: {
    required: ["name"],
    validate: (op) => typeof op.name === "string" && op.name.length > 0,
  },
  setText: {
    required: ["slide", "shape", "text"],
    validate: (op) =>
      typeof op.slide === "string" && typeof op.shape === "string" && typeof op.text === "string",
  },
  addShape: {
    required: ["slide", "shapeType", "name"],
    validate: (op) =>
      typeof op.slide === "string" && typeof op.shapeType === "string" && typeof op.name === "string",
  },
  deleteShape: {
    required: ["slide", "name"],
    validate: (op) =>
      typeof op.slide === "string" && typeof op.name === "string",
  },
  addTable: {
    required: ["slide", "rows", "columns", "left", "top"],
    validate: (op) =>
      typeof op.slide === "string" &&
      typeof op.rows === "number" &&
      typeof op.columns === "number" &&
      typeof op.left === "number" &&
      typeof op.top === "number",
  },
  addChart: {
    required: ["slide", "chartType", "data"],
    validate: (op) =>
      typeof op.slide === "string" && typeof op.chartType === "string" && Array.isArray(op.data),
  },
  readShapes: {
    required: ["slide"],
    validate: (op) => typeof op.slide === "string",
  },
};

/**
 * Execute a single PowerPoint operation within context.
 */
export async function executePptOperation(context, op) {
  const presentation = context.presentation;
  const type = op.operation;

  switch (type) {
    case "addSlide": {
      console.log(`[ppt:op:addSlide] Creating slide "${op.name}"`);
      try {
        const slides = presentation.slides;
        slides.load("items/name");
        await context.sync();

        const existing = slides.items.some((s) => s.name === op.name);
        if (existing) {
          console.log(`[ppt:op:addSlide] Slide "${op.name}" already exists`);
          return `Slide "${op.name}" already exists`;
        }

        const slide = slides.add(PowerPoint.SlideLayoutType.BLANK);
        slide.name = op.name;
        await context.sync();
        console.log(`[ppt:op:addSlide] Slide "${op.name}" ready`);
        return `Slide "${op.name}" ready`;
      } catch (e) {
        console.error(`[ppt:op:addSlide] Failed for "${op.name}":`, e.message);
        throw e;
      }
    }

    case "deleteSlide": {
      console.log(`[ppt:op:deleteSlide] Deleting slide "${op.name}"`);
      try {
        const slides = presentation.slides;
        slides.load("items/name");
        await context.sync();

        const slide = slides.getItemOrNullObject(op.name);
        slide.load("isNullObject");
        await context.sync();

        if (slide.isNullObject) {
          const allNames = slides.items.map((s) => s.name).join(", ");
          return `Skipped delete: slide "${op.name}" not found. Available: ${allNames}`;
        }

        slide.delete();
        await context.sync();
        console.log(`[ppt:op:deleteSlide] Deleted slide "${op.name}"`);
        return `Deleted slide "${op.name}"`;
      } catch (e) {
        console.error(`[ppt:op:deleteSlide] Failed:`, e.message);
        throw e;
      }
    }

    case "setText": {
      console.log(`[ppt:op:setText] slide="${op.slide}" shape="${op.shape}"`);
      try {
        const slides = presentation.slides;
        slides.load("items/name");
        await context.sync();

        const slide = slides.getItem(op.slide);
        slide.load("name", "shapes");
        await context.sync();

        const shapes = slide.shapes;
        shapes.load("items/name");
        await context.sync();

        const shape = shapes.getItemOrNullObject(op.shape);
        shape.load("isNullObject");
        await context.sync();

        if (shape.isNullObject) {
          const available = shapes.items.map((s) => s.name).join(", ");
          throw new Error(
            `Shape "${op.shape}" not found on slide "${op.slide}". Available: [${available}]`
          );
        }

        shape.text = op.text;
        await context.sync();
        console.log(`[ppt:op:setText] Set text on ${op.slide}/${op.shape}`);
        return `Set text on ${op.slide}/${op.shape}`;
      } catch (e) {
        console.error(`[ppt:op:setText] Failed:`, e.message);
        throw e;
      }
    }

    case "addShape": {
      console.log(`[ppt:op:addShape] slide="${op.slide}" type="${op.shapeType}" name="${op.name}"`);
      try {
        const slides = presentation.slides;
        slides.load("items/name");
        await context.sync();

        const slide = slides.getItem(op.slide);
        slide.load("name");
        await context.sync();

        // slide.shapes accessed via addShape - no load needed for adding
        const shapeTypeMap = {
          rectangle: PowerPoint.ShapeType.rectangle,
          ellipse: PowerPoint.ShapeType.ellipse,
          triangle: PowerPoint.ShapeType.triangle,
          line: PowerPoint.ShapeType.line,
          image: PowerPoint.ShapeType.image,
          textbox: PowerPoint.ShapeType.textCloud,
        };

        const shapeTypeEnum = shapeTypeMap[op.shapeType] || PowerPoint.ShapeType.rectangle;
        const shape = slide.shapes.addShape(shapeTypeEnum, op.left || 50, op.top || 50, op.width || 200, op.height || 100);
        shape.name = op.name;
        if (op.text) {
          shape.text = op.text;
        }
        await context.sync();
        console.log(`[ppt:op:addShape] Added shape "${op.name}" on ${op.slide}`);
        return `Added shape "${op.name}" on ${op.slide}`;
      } catch (e) {
        console.error(`[ppt:op:addShape] Failed:`, e.message);
        throw e;
      }
    }

    case "deleteShape": {
      console.log(`[ppt:op:deleteShape] slide="${op.slide}" shape="${op.name}"`);
      try {
        const slides = presentation.slides;
        slides.load("items/name");
        await context.sync();

        const slide = slides.getItem(op.slide);
        slide.load("name", "shapes");
        await context.sync();

        const shapes = slide.shapes;
        shapes.load("items/name");
        await context.sync();

        const shape = shapes.getItemOrNullObject(op.name);
        shape.load("isNullObject");
        await context.sync();

        if (shape.isNullObject) {
          const available = shapes.items.map((s) => s.name).join(", ");
          return `Skipped delete: shape "${op.name}" not found. Available: [${available}]`;
        }

        shape.delete();
        await context.sync();
        console.log(`[ppt:op:deleteShape] Deleted shape "${op.name}"`);
        return `Deleted shape "${op.name}" on ${op.slide}`;
      } catch (e) {
        console.error(`[ppt:op:deleteShape] Failed:`, e.message);
        throw e;
      }
    }

    case "addTable": {
      console.log(`[ppt:op:addTable] slide="${op.slide}" ${op.rows}x${op.columns}`);
      try {
        const slides = presentation.slides;
        slides.load("items/name");
        await context.sync();

        const slide = slides.getItem(op.slide);
        slide.load("name", "tables");
        await context.sync();

        slide.tables.add(op.rows, op.columns, op.top, op.left, op.hidden ?? false);
        await context.sync();
        console.log(`[ppt:op:addTable] Added table on ${op.slide}`);
        return `Added ${op.rows}x${op.columns} table on ${op.slide}`;
      } catch (e) {
        console.error(`[ppt:op:addTable] Failed:`, e.message);
        throw e;
      }
    }

    case "addChart": {
      console.log(`[ppt:op:addChart] slide="${op.slide}" type="${op.chartType}"`);
      try {
        const slides = presentation.slides;
        slides.load("items/name");
        await context.sync();

        const slide = slides.getItem(op.slide);
        slide.load("name", "charts");
        await context.sync();

        const chartTypeMap = {
          column: PowerPoint.ChartType.columnClustered,
          bar: PowerPoint.ChartType.barClustered,
          line: PowerPoint.ChartType.line,
          pie: PowerPoint.ChartType.pie,
          area: PowerPoint.ChartType.areaStacked,
        };

        const chartTypeEnum = chartTypeMap[op.chartType] || PowerPoint.ChartType.columnClustered;
        const chart = slide.charts.add(chartTypeEnum, op.data);
        chart.chartType = chartTypeEnum;
        await context.sync();
        console.log(`[ppt:op:addChart] Added chart on ${op.slide}`);
        return `Added chart "${op.chartType}" on ${op.slide}`;
      } catch (e) {
        console.error(`[ppt:op:addChart] Failed:`, e.message);
        throw e;
      }
    }

    case "readShapes": {
      console.log(`[ppt:op:readShapes] slide="${op.slide}"`);
      try {
        const slides = presentation.slides;
        slides.load("items/name");
        await context.sync();

        const slide = slides.getItem(op.slide);
        slide.load("name", "shapes");
        await context.sync();

        const shapes = slide.shapes;
        shapes.load("items/name, text, type");
        await context.sync();

        const shapeList = shapes.items.map((shape) => {
          const textPreview = shape.text ? shape.text.substring(0, 80).replace(/\n/g, " ") : "(no text)";
          return `  ${shape.name} (${shape.type}): "${textPreview}"`;
        });

        return `${shapes.items.length} shape(s) on ${slide.name}:\n${shapeList.join("\n")}`;
      } catch (e) {
        console.error(`[ppt:op:readShapes] Failed:`, e.message);
        throw e;
      }
    }

    default:
      throw new Error(`Unknown PowerPoint operation type: "${type}"`);
  }
}
