/* global Excel, console */

/* Operation registry and Excel execution engine */

/* Excel-specific operation registry */
export const EXCEL_OPERATION_REGISTRY = {
  // === Worksheet Management ===
  createWorksheet: {
    required: ["name"],
    validate: (op) => typeof op.name === "string" && op.name.length > 0,
  },
  deleteWorksheet: {
    required: ["name"],
    validate: (op) => typeof op.name === "string" && op.name.length > 0,
  },
  renameWorksheet: {
    required: ["name", "newName"],
    validate: (op) => typeof op.name === "string" && typeof op.newName === "string",
  },
  activateWorksheet: {
    required: ["name"],
    validate: (op) => typeof op.name === "string",
  },
  copyWorksheet: {
    required: ["name", "targetSheet", "after"],
    validate: (op) => typeof op.name === "string" && typeof op.targetSheet === "string",
  },
  hideWorksheet: {
    required: ["name"],
    validate: (op) => typeof op.name === "string",
  },
  showWorksheet: {
    required: ["name"],
    validate: (op) => typeof op.name === "string",
  },

  // === Data Operations ===
  writeValues: {
    required: ["sheet", "range", "values"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && Array.isArray(op.values),
  },
  writeFormulas: {
    required: ["sheet", "range", "formulas"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && Array.isArray(op.formulas),
  },
  readRange: {
    required: ["sheet", "range"],
    validate: (op) => typeof op.sheet === "string" && typeof op.range === "string",
  },

  // === Cell Formatting ===
  setNumberFormat: {
    required: ["sheet", "range", "format"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.format === "string",
  },
  setFillColor: {
    required: ["sheet", "range", "color"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.color === "string",
  },
  setFontColor: {
    required: ["sheet", "range", "color"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.color === "string",
  },
  setFontBold: {
    required: ["sheet", "range", "bold"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.bold === "boolean",
  },
  setFontItalic: {
    required: ["sheet", "range", "italic"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.range === "string" &&
      typeof op.italic === "boolean",
  },
  setFontUnderline: {
    required: ["sheet", "range", "underline"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.range === "string" &&
      typeof op.underline === "boolean",
  },
  setFontStrikethrough: {
    required: ["sheet", "range", "strikethrough"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.range === "string" &&
      typeof op.strikethrough === "boolean",
  },
  setFontName: {
    required: ["sheet", "range", "font"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.font === "string",
  },
  setFontSize: {
    required: ["sheet", "range", "size"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.size === "number",
  },
  setHorizontalAlignment: {
    required: ["sheet", "range", "alignment"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.range === "string" &&
      typeof op.alignment === "string",
  },
  setVerticalAlignment: {
    required: ["sheet", "range", "alignment"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.range === "string" &&
      typeof op.alignment === "string",
  },
  setTextWrap: {
    required: ["sheet", "range", "wrap"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.wrap === "boolean",
  },
  setIndentLevel: {
    required: ["sheet", "range", "level"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.level === "number",
  },
  setRotationAngle: {
    required: ["sheet", "range", "angle"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.angle === "number",
  },
  setBorder: {
    required: ["sheet", "range", "location", "style", "color"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.range === "string" &&
      typeof op.location === "string" &&
      typeof op.style === "string" &&
      typeof op.color === "string",
  },
  setMergeCells: {
    required: ["sheet", "range", "merge"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.merge === "boolean",
  },
  setHyperlink: {
    required: ["sheet", "range", "link"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.link === "string",
  },
  setCellProtection: {
    required: ["sheet", "range", "locked", "hidden"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.range === "string" &&
      typeof op.locked === "boolean",
  },

  // === Autofit ===
  autofitColumns: {
    required: ["sheet"],
    validate: (op) => typeof op.sheet === "string",
  },
  autofitRows: {
    required: ["sheet"],
    validate: (op) => typeof op.sheet === "string",
  },

  // === Tables ===
  createTable: {
    required: ["sheet", "range", "name"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.name === "string",
  },
  deleteTable: {
    required: ["name"],
    validate: (op) => typeof op.name === "string",
  },
  addTableRow: {
    required: ["tableName", "index", "values"],
    validate: (op) => typeof op.tableName === "string" && Array.isArray(op.values),
  },
  addTableColumn: {
    required: ["tableName", "name"],
    validate: (op) => typeof op.tableName === "string" && typeof op.name === "string",
  },
  deleteTableRow: {
    required: ["tableName", "index"],
    validate: (op) => typeof op.tableName === "string" && typeof op.index === "number",
  },
  deleteTableColumn: {
    required: ["tableName", "index"],
    validate: (op) => typeof op.tableName === "string" && typeof op.index === "number",
  },
  sortTable: {
    required: ["tableName", "columns"],
    validate: (op) => typeof op.tableName === "string" && Array.isArray(op.columns),
  },
  filterTable: {
    required: ["tableName", "columnIndex", "rule"],
    validate: (op) => typeof op.tableName === "string" && typeof op.columnIndex === "number",
  },
  setTableStyle: {
    required: ["tableName", "style"],
    validate: (op) => typeof op.tableName === "string" && typeof op.style === "string",
  },
  setTableShowHeader: {
    required: ["tableName", "show"],
    validate: (op) => typeof op.tableName === "string" && typeof op.show === "boolean",
  },
  setTableShowTotals: {
    required: ["tableName", "show"],
    validate: (op) => typeof op.tableName === "string" && typeof op.show === "boolean",
  },
  setTableBandedRows: {
    required: ["tableName", "banded"],
    validate: (op) => typeof op.tableName === "string" && typeof op.banded === "boolean",
  },
  setTableBandedColumns: {
    required: ["tableName", "banded"],
    validate: (op) => typeof op.tableName === "string" && typeof op.banded === "boolean",
  },

  // === Charts ===
  addChart: {
    required: ["sheet", "dataRange", "chartType", "title"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.dataRange === "string" &&
      typeof op.chartType === "string" &&
      typeof op.title === "string",
  },
  deleteChart: {
    required: ["sheet", "name"],
    validate: (op) => typeof op.sheet === "string" && typeof op.name === "string",
  },
  setChartTitle: {
    required: ["sheet", "name", "title"],
    validate: (op) => typeof op.sheet === "string" && typeof op.name === "string",
  },
  setChartSeries: {
    required: ["sheet", "name", "dataRange", "seriesName"],
    validate: (op) => typeof op.sheet === "string" && typeof op.name === "string",
  },
  setChartPosition: {
    required: ["sheet", "name", "left", "top", "width", "height"],
    validate: (op) => typeof op.sheet === "string" && typeof op.name === "string",
  },
  setChartShowLegend: {
    required: ["sheet", "name", "show"],
    validate: (op) => typeof op.sheet === "string" && typeof op.name === "string",
  },
  setChartShowDataLabels: {
    required: ["sheet", "name", "show"],
    validate: (op) => typeof op.sheet === "string" && typeof op.name === "string",
  },

  // === PivotTables ===
  createPivotTable: {
    required: ["name", "sourceData", "pivotSheet", "pivotRange"],
    validate: (op) =>
      typeof op.name === "string" &&
      typeof op.sourceData === "string" &&
      typeof op.pivotSheet === "string" &&
      typeof op.pivotRange === "string",
  },
  addPivotHierarchy: {
    required: ["pivotTable", "field", "category", "index"],
    validate: (op) =>
      typeof op.pivotTable === "string" &&
      typeof op.field === "string" &&
      typeof op.category === "string",
  },
  setPivotAggregation: {
    required: ["pivotTable", "field", "aggregation"],
    validate: (op) =>
      typeof op.pivotTable === "string" &&
      typeof op.field === "string" &&
      typeof op.aggregation === "string",
  },
  setPivotLayout: {
    required: ["pivotTable", "layout"],
    validate: (op) => typeof op.pivotTable === "string" && typeof op.layout === "string",
  },
  addPivotFilter: {
    required: ["pivotTable", "field", "criterion"],
    validate: (op) => typeof op.pivotTable === "string" && typeof op.field === "string",
  },
  deletePivotTable: {
    required: ["name"],
    validate: (op) => typeof op.name === "string",
  },

  // === Conditional Formatting ===
  addDataBar: {
    required: ["sheet", "range", "color", "minType", "minValue", "maxType", "maxValue"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.color === "string",
  },
  addColorScale: {
    required: ["sheet", "range", "colors"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && Array.isArray(op.colors),
  },
  addIconSet: {
    required: ["sheet", "range", "iconSet"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.range === "string" &&
      typeof op.iconSet === "string",
  },
  addTextComparison: {
    required: ["sheet", "range", "textComparison", "text", "operator"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.range === "string" &&
      typeof op.textComparison === "string" &&
      typeof op.operator === "string",
  },
  addTopBottom: {
    required: ["sheet", "range", "topBottom", "value", "belowAverage"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.range === "string" &&
      typeof op.topBottom === "number",
  },
  addPresetCriteria: {
    required: ["sheet", "range", "criteria"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.range === "string" &&
      typeof op.criteria === "string",
  },
  addCustomConditionalFormat: {
    required: ["sheet", "range", "formula"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.range === "string" &&
      typeof op.formula === "string",
  },
  clearConditionalFormats: {
    required: ["sheet", "range"],
    validate: (op) => typeof op.sheet === "string" && typeof op.range === "string",
  },

  // === Shapes ===
  addShape: {
    required: ["sheet", "shapeType", "name", "left", "top", "width", "height"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.shapeType === "string" &&
      typeof op.name === "string" &&
      typeof op.left === "number" &&
      typeof op.top === "number",
  },
  addTextBox: {
    required: ["sheet", "name", "left", "top", "width", "height", "text"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.name === "string" &&
      typeof op.left === "number" &&
      typeof op.top === "number",
  },
  setShapeText: {
    required: ["sheet", "name", "text"],
    validate: (op) => typeof op.sheet === "string" && typeof op.name === "string",
  },
  deleteShape: {
    required: ["sheet", "name"],
    validate: (op) => typeof op.sheet === "string" && typeof op.name === "string",
  },

  // === Data Validation ===
  addDataValidation: {
    required: ["sheet", "range", "ruleType", "operator", "value1"],
    validate: (op) =>
      typeof op.sheet === "string" &&
      typeof op.range === "string" &&
      typeof op.ruleType === "string",
  },
  deleteDataValidation: {
    required: ["sheet", "range"],
    validate: (op) => typeof op.sheet === "string" && typeof op.range === "string",
  },

  // === Comments & Notes ===
  addComment: {
    required: ["sheet", "range", "text"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.text === "string",
  },
  deleteComment: {
    required: ["sheet", "range"],
    validate: (op) => typeof op.sheet === "string" && typeof op.range === "string",
  },
  addNote: {
    required: ["sheet", "range", "text"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.text === "string",
  },

  // === Named Items ===
  addNamedRange: {
    required: ["name", "definition"],
    validate: (op) => typeof op.name === "string" && typeof op.definition === "string",
  },
  deleteNamedRange: {
    required: ["name"],
    validate: (op) => typeof op.name === "string",
  },

  // === Worksheet View & Layout ===
  setFreezePanes: {
    required: ["sheet", "freezeType", "row", "column"],
    validate: (op) => typeof op.sheet === "string" && typeof op.freezeType === "string",
  },
  setPageLayout: {
    required: ["sheet", "orientation", "pageSize", "margins"],
    validate: (op) => typeof op.sheet === "string",
  },
  setPrintSettings: {
    required: ["sheet", "printArea", "printTitles"],
    validate: (op) => typeof op.sheet === "string",
  },
  addPageBreak: {
    required: ["sheet", "position"],
    validate: (op) => typeof op.sheet === "string" && typeof op.position === "string",
  },

  // === Protection ===
  protectWorksheet: {
    required: ["name", "password"],
    validate: (op) => typeof op.name === "string",
  },
  unprotectWorksheet: {
    required: ["name", "password"],
    validate: (op) => typeof op.name === "string",
  },

  // === Sort ===
  sortRange: {
    required: ["sheet", "range", "key", "order"],
    validate: (op) =>
      typeof op.sheet === "string" && typeof op.range === "string" && typeof op.key === "number",
  },

  // === Delete ===
     deleteRows: {
       required: ["sheet", "range"],
       validate: (op) => typeof op.sheet === "string" && typeof op.range === "string",
     },
     deleteColumns: {
       required: ["sheet", "range"],
       validate: (op) => typeof op.sheet === "string" && typeof op.range === "string",
     },

     // === Clear Sheet ===
     clearSheet: {
       required: ["sheet"],
       validate: (op) => typeof op.sheet === "string",
     },

     // === Slicers ===
  addSlicer: {
    required: ["pivotTable", "field", "slicerName", "sheet"],
    validate: (op) =>
      typeof op.pivotTable === "string" &&
      typeof op.field === "string" &&
      typeof op.slicerName === "string",
  },
  deleteSlicer: {
    required: ["sheet", "name"],
    validate: (op) => typeof op.sheet === "string" && typeof op.name === "string",
  },
};

/**
 * Map chart type strings to Excel.ChartType enum values.
 * Wrapped in try/catch — ChartType may not exist in older Excel versions.
 */
let _chartTypeMap = null;
export function getChartTypeMap() {
  if (!_chartTypeMap) {
    try {
      _chartTypeMap = {
        line: Excel.ChartType.line,
        column: Excel.ChartType.columnClustered,
        columnStacked: Excel.ChartType.columnStacked,
        columnStacked100: Excel.ChartType.columnStacked100,
        bar: Excel.ChartType.barClustered,
        barStacked: Excel.ChartType.barStacked,
        barStacked100: Excel.ChartType.barStacked100,
        pie: Excel.ChartType.pie,
        pieOfPie: Excel.ChartType.pieOfPie,
        doughnut: Excel.ChartType.doughnut,
        area: Excel.ChartType.areaStacked,
        areaStacked100: Excel.ChartType.areaStacked100,
        xyScatter: Excel.ChartType.xyscatter,
        xyScatterSmooth: Excel.ChartType.xyscatterSmooth,
        xyScatterLines: Excel.ChartType.xyscatterLines,
        radar: Excel.ChartType.radar,
        radarFilled: Excel.ChartType.radarFilled,
        surface: Excel.ChartType.surface,
        surfaceWireframe: Excel.ChartType.surfaceWireframe,
        stockOpenHighLowClose: Excel.ChartType.stockOpenHighLowClose,
        volume: Excel.ChartType.volume,
      };
    } catch (e) {
      console.warn('[operations] ChartType enum not available:', e.message);
      _chartTypeMap = {};
    }
  }
  return _chartTypeMap;
}

/**
 * Lazily get pivot table aggregation enum map.
 * Pivot table API may not be available in all Excel versions.
 */
let _pivotAggregationMap = null;
function getPivotAggregationMap() {
  if (!_pivotAggregationMap && Excel.PivotDataAggregation) {
    _pivotAggregationMap = {
      sum: Excel.PivotDataAggregation.sum,
      count: Excel.PivotDataAggregation.count,
      average: Excel.PivotDataAggregation.average,
      max: Excel.PivotDataAggregation.max,
      min: Excel.PivotDataAggregation.min,
      product: Excel.PivotDataAggregation.product,
      countNumbers: Excel.PivotDataAggregation.countNumbers,
      standardDeviation: Excel.PivotDataAggregation.standardDeviation,
      standardDeviationP: Excel.PivotDataAggregation.standardDeviationP,
      variance: Excel.PivotDataAggregation.variance,
      varianceP: Excel.PivotDataAggregation.varianceP,
    };
  }
  return _pivotAggregationMap;
}

/**
 * Lazily get pivot table layout enum map.
 * Pivot table API may not be available in all Excel versions.
 */
let _pivotLayoutMap = null;
function getPivotLayoutMap() {
  if (!_pivotLayoutMap && Excel.PivotLayoutType) {
    _pivotLayoutMap = {
      compact: Excel.PivotLayoutType.compact,
      outline: Excel.PivotLayoutType.outline,
      tabular: Excel.PivotLayoutType.tabular,
    };
  }
  return _pivotLayoutMap;
}

/**
 * Lazily get pivot hierarchy insert position enum.
 */
let _pivotHierarchyInsertPosition = null;
function getPivotHierarchyInsertPosition() {
  if (!_pivotHierarchyInsertPosition && Excel.PivotHierarchyInsertPosition) {
    _pivotHierarchyInsertPosition = Excel.PivotHierarchyInsertPosition;
  }
  return _pivotHierarchyInsertPosition;
}

/**
 * Map conditional formatting icon sets.
 * Wrapped in try/catch — IconSetType may not exist in older Excel versions.
 */
let _iconSetsMap = null;
function getIconSetsMap() {
  if (!_iconSetsMap) {
    try {
      _iconSetsMap = {
        fiveArrows: Excel.IconSetType.fiveArrows,
        fiveArrowsGray: Excel.IconSetType.fiveArrowsGray,
        fiveBoxes: Excel.IconSetType.fiveBoxes,
        fiveQuarters: Excel.IconSetType.fiveQuarters,
        fiveRating: Excel.IconSetType.fiveRating,
        fourArrows: Excel.IconSetType.fourArrows,
        fourArrowsGray: Excel.IconSetType.fourArrowsGray,
        fourRating: Excel.IconSetType.fourRating,
        fourRedToBlack: Excel.IconSetType.fourRedToBlack,
        fourTrafficLights: Excel.IconSetType.fourTrafficLights,
        threeArrows: Excel.IconSetType.threeArrows,
        threeFlags: Excel.IconSetType.threeFlags,
        threeSigns: Excel.IconSetType.threeSigns,
        threeSymbols: Excel.IconSetType.threeSymbols,
        threeSymbols2: Excel.IconSetType.threeSymbols2,
        threeTrafficLights1: Excel.IconSetType.threeTrafficLights1,
        threeTrafficLights2: Excel.IconSetType.threeTrafficLights2,
      };
    } catch (e) {
      console.warn('[operations] IconSetType enum not available:', e.message);
      _iconSetsMap = {};
    }
  }
  return _iconSetsMap;
}

/**
 * Map text comparison operators.
 */
let _textComparisonMap = null;
function getTextComparisonMap() {
  if (!_textComparisonMap) {
    try {
      _textComparisonMap = {
        containsText: Excel.TextComparisonOperator.containsText,
        doesNotContain: Excel.TextComparisonOperator.doesNotContain,
        beginsWith: Excel.TextComparisonOperator.beginsWith,
        endsWith: Excel.TextComparisonOperator.endsWith,
      };
    } catch (e) {
      console.warn('[operations] TextComparisonOperator enum not available:', e.message);
      _textComparisonMap = {};
    }
  }
  return _textComparisonMap;
}

/**
 * Map preset criteria.
 */
let _presetCriteriaMap = null;
function getPresetCriteriaMap() {
  if (!_presetCriteriaMap) {
    try {
      _presetCriteriaMap = {
        aboveAverage: Excel.PresetCriteria.aboveAverage,
        belowAverage: Excel.PresetCriteria.belowAverage,
        aboveOrEqualAverage: Excel.PresetCriteria.aboveOrEqualAverage,
        belowOrEqualAverage: Excel.PresetCriteria.belowOrEqualAverage,
        aboveStdDev1: Excel.PresetCriteria.aboveStdDev1,
        aboveStdDev2: Excel.PresetCriteria.aboveStdDev2,
        aboveStdDev3: Excel.PresetCriteria.aboveStdDev3,
        belowStdDev1: Excel.PresetCriteria.belowStdDev1,
        belowStdDev2: Excel.PresetCriteria.belowStdDev2,
        belowStdDev3: Excel.PresetCriteria.belowStdDev3,
        uniqueValues: Excel.PresetCriteria.uniqueValues,
        duplicateValues: Excel.PresetCriteria.duplicateValues,
        containsBlank: Excel.PresetCriteria.containsBlank,
        notContainsBlank: Excel.PresetCriteria.notContainsBlank,
        containsErrors: Excel.PresetCriteria.containsErrors,
        notContainsErrors: Excel.PresetCriteria.notContainsErrors,
      };
    } catch (e) {
      console.warn('[operations] PresetCriteria enum not available:', e.message);
      _presetCriteriaMap = {};
    }
  }
  return _presetCriteriaMap;
}

/**
 * Map top/bottom rule type strings.
 */
const TOP_BOTTOM_MAP = {
  top: 0,
  bottom: 1,
};

/**
 * Map data validation rule type strings.
 */
let _validationRuleMap = null;
function getValidationRuleMap() {
  if (!_validationRuleMap) {
    try {
      _validationRuleMap = {
        wholeNumber: Excel.DataValidationType.wholeNumber,
        decimal: Excel.DataValidationType.decimal,
        date: Excel.DataValidationType.date,
        time: Excel.DataValidationType.time,
        textLength: Excel.DataValidationType.textLength,
        list: Excel.DataValidationType.list,
        custom: Excel.DataValidationType.custom,
      };
    } catch (e) {
      console.warn('[operations] DataValidationType enum not available:', e.message);
      _validationRuleMap = {};
    }
  }
  return _validationRuleMap;
}

/**
 * Map data validation operator strings.
 */
let _validationOperatorMap = null;
function getValidationOperatorMap() {
  if (!_validationOperatorMap) {
    try {
      _validationOperatorMap = {
        greaterThan: Excel.DataValidationOperator.greaterThan,
        lessThan: Excel.DataValidationOperator.lessThan,
        between: Excel.DataValidationOperator.between,
        notBetween: Excel.DataValidationOperator.notBetween,
        equal: Excel.DataValidationOperator.equal,
        notEqual: Excel.DataValidationOperator.notEqual,
      };
    } catch (e) {
      console.warn('[operations] DataValidationOperator enum not available:', e.message);
      _validationOperatorMap = {};
    }
  }
  return _validationOperatorMap;
}

/**
 * Map sort order strings.
 */
let _sortOrderMap = null;
function getSortOrderMap() {
  if (!_sortOrderMap) {
    try {
      _sortOrderMap = {
        ascending: Excel.SortOrder.ascending,
        descending: Excel.SortOrder.descending,
      };
    } catch (e) {
      console.warn('[operations] SortOrder enum not available:', e.message);
      _sortOrderMap = { ascending: 0, descending: 1 };
    }
  }
  return _sortOrderMap;
}

/**
 * Map horizontal alignment strings.
 */
let _hAlignMap = null;
function getHAlignMap() {
  if (!_hAlignMap) {
    try {
      _hAlignMap = {
        left: Excel.HorizontalAlignment.left,
        center: Excel.HorizontalAlignment.center,
        right: Excel.HorizontalAlignment.right,
        justify: Excel.HorizontalAlignment.justify,
        fill: Excel.HorizontalAlignment.fill,
        general: Excel.HorizontalAlignment.general,
        centerAcrossSelection: Excel.HorizontalAlignment.centerAcrossSelection,
        distributed: Excel.HorizontalAlignment.distributed,
      };
    } catch (e) {
      console.warn('[operations] HorizontalAlignment enum not available:', e.message);
      _hAlignMap = {};
    }
  }
  return _hAlignMap;
}

/**
 * Map vertical alignment strings.
 */
let _vAlignMap = null;
function getVAlignMap() {
  if (!_vAlignMap) {
    try {
      _vAlignMap = {
        top: Excel.VerticalAlignment.top,
        center: Excel.VerticalAlignment.center,
        bottom: Excel.VerticalAlignment.bottom,
        justify: Excel.VerticalAlignment.justify,
        distributed: Excel.VerticalAlignment.distributed,
      };
    } catch (e) {
      console.warn('[operations] VerticalAlignment enum not available:', e.message);
      _vAlignMap = {};
    }
  }
  return _vAlignMap;
}

/**
 * Map border location strings.
 */
let _borderLocationMap = null;
function getBorderLocationMap() {
  if (!_borderLocationMap) {
    try {
      _borderLocationMap = {
        edgeTop: Excel.BorderLocation.edgeTop,
        edgeBottom: Excel.BorderLocation.edgeBottom,
        edgeLeft: Excel.BorderLocation.edgeLeft,
        edgeRight: Excel.BorderLocation.edgeRight,
        insideHorizontal: Excel.BorderLocation.insideHorizontal,
        insideVertical: Excel.BorderLocation.insideVertical,
        diagonalUp: Excel.BorderLocation.diagonalUp,
        diagonalDown: Excel.BorderLocation.diagonalDown,
      };
    } catch (e) {
      console.warn('[operations] BorderLocation enum not available:', e.message);
      _borderLocationMap = {};
    }
  }
  return _borderLocationMap;
}

/**
 * Map border style strings.
 */
let _borderStyleMap = null;
function getBorderStyleMap() {
  if (!_borderStyleMap) {
    try {
      _borderStyleMap = {
        none: Excel.BorderStyle.none,
        continuous: Excel.BorderStyle.continuous,
        dash: Excel.BorderStyle.dash,
        dot: Excel.BorderStyle.dot,
        double: Excel.BorderStyle.double,
        dashDot: Excel.BorderStyle.dashDot,
        dashDotDot: Excel.BorderStyle.dashDotDot,
        slantDashDot: Excel.BorderStyle.slantDashDot,
      };
    } catch (e) {
      console.warn('[operations] BorderStyle enum not available:', e.message);
      _borderStyleMap = { continuous: 0 };
    }
  }
  return _borderStyleMap;
}

/**
 * Auto-adjust a range string to match actual data dimensions.
 * E.g. "A1:F1" with 1×1 data → "A1:A1"
 * E.g. "A1:Z100" with 3×5 data → "A1:E3"
 * Preserves the top-left anchor of the original range.
 */
export function adjustRangeToData(sheet, rangeStr, dataRows, dataCols) {
  const match = rangeStr.match(/^([A-Z]+)(\d+)/i);
  if (!match) {
    console.warn(`[adjustRangeToData] Could not parse range "${rangeStr}", using as-is`);
    return rangeStr;
  }
  const colLetter = match[1].toUpperCase();
  const rowNum = parseInt(match[2], 10);
  let colIndex = 0;
  for (let i = 0; i < colLetter.length; i++) {
    colIndex = colIndex * 26 + (colLetter.charCodeAt(i) - 64);
  }
  const startRow = rowNum - 1;
  const startCol = colIndex - 1;
  const endRow = startRow + dataRows;
  const endCol = startCol + dataCols;

  function colToLetter(col) {
    let letter = "";
    while (col >= 0) {
      letter = String.fromCharCode((col % 26) + 65) + letter;
      col = Math.floor(col / 26) - 1;
    }
    return letter;
  }
  return `${colToLetter(startCol)}${startRow + 1}:${colToLetter(endCol - 1)}${endRow}`;
}

/**
 * Execute a single operation within Excel context.
 * @param {Excel.RequestContext} context
 * @param {object} op
 * @returns {string} Result message
 */
export async function executeOperation(context, op) {
  const workbook = context.workbook;
  const type = op.operation;

  switch (type) {
    // === Worksheet Management ===
    case "createWorksheet": {
      console.log(`[op:createWorksheet] Creating sheet "${op.name}"`);
      try {
        const sheet = workbook.worksheets.getItemOrNullObject(op.name);
        sheet.load("exists");
        await context.sync();
        if (!sheet.exists) {
          workbook.worksheets.add(op.name);
          await context.sync();
        }
        console.log(`[op:createWorksheet] Sheet "${op.name}" ready`);
        return `Sheet "${op.name}" ready`;
      } catch (e) {
        console.error(`[op:createWorksheet] Failed for "${op.name}":`, e.message);
        throw e;
      }
    }

    case "deleteWorksheet": {
      const sheet9 = workbook.worksheets.getItemOrNullObject(op.name);
      sheet9.load("name, isNullObject");
      await context.sync();
      if (sheet9.isNullObject) {
        const allSheets = workbook.worksheets;
        allSheets.load("items/name");
        await context.sync();
        const sheetNames = allSheets.items.map((s) => s.name);
        console.warn(
          `[deleteWorksheet] Sheet "${op.name}" does not exist, available: ${sheetNames.join(", ")}`
        );
        return `Skipped delete: sheet "${op.name}" not found. Available sheets: ${sheetNames.join(", ")}`;
      }
      console.log(`[deleteWorksheet] Deleting sheet "${sheet9.name}"`);
      sheet9.delete();
      await context.sync();
      return `Deleted sheet "${op.name}"`;
    }

    case "renameWorksheet": {
      const sheet = workbook.worksheets.getItemOrNullObject(op.name);
      sheet.load("name, isNullObject");
      await context.sync();
      if (sheet.isNullObject) {
        throw new Error(`Sheet "${op.name}" not found.`);
      }
      sheet.name = op.newName;
      await context.sync();
      console.log(`[op:renameWorksheet] "${op.name}" → "${op.newName}"`);
      return `Renamed sheet "${op.name}" → "${op.newName}"`;
    }

    case "activateWorksheet": {
      const sheet = workbook.worksheets.getItem(op.name);
      sheet.activate();
      await context.sync();
      console.log(`[op:activateWorksheet] Activated "${op.name}"`);
      return `Activated sheet "${op.name}"`;
    }

    case "copyWorksheet": {
      const source = workbook.worksheets.getItem(op.name);
      source.load("name");
      await context.sync();
      const target = workbook.worksheets.getItem(op.targetSheet);
      target.load("name");
      await context.sync();
      const insertPos = op.after
         ? (Excel.WorksheetInsertPosition ? Excel.WorksheetInsertPosition.after : 1)
         : (Excel.WorksheetInsertPosition ? Excel.WorksheetInsertPosition.before : 0);
       source.copyInto(op.targetSheet, insertPos);
      await context.sync();
      console.log(
        `[op:copyWorksheet] "${op.name}" → "${op.targetSheet}" (${op.after ? "after" : "before"})`
      );
      return `Copied sheet "${op.name}" into "${op.targetSheet}"`;
    }

    case "hideWorksheet": {
      const sheet = workbook.worksheets.getItem(op.name);
      sheet.visibility = Excel.Visibility ? Excel.Visibility.hidden : 1;
      await context.sync();
      console.log(`[op:hideWorksheet] Hidden "${op.name}"`);
      return `Hidden sheet "${op.name}"`;
    }

    case "showWorksheet": {
      const sheet = workbook.worksheets.getItem(op.name);
      sheet.visibility = Excel.Visibility ? Excel.Visibility.visible : 0;
      await context.sync();
      console.log(`[op:showWorksheet] Shown "${op.name}"`);
      return `Shown sheet "${op.name}"`;
    }

    // === Data Operations ===
    case "writeValues": {
      console.log(
        `[op:writeValues] sheet="${op.sheet}" range="${op.range}" rows=${op.values.length}`
      );
      console.log(`[op:writeValues] sample data:`, JSON.stringify(op.values.slice(0, 2)));
      try {
        let sheet;
        try {
          sheet = workbook.worksheets.getItem(op.sheet);
          sheet.load("name");
          await context.sync();
          console.log(`[op:writeValues] Sheet "${op.sheet}" found`);
        } catch {
          console.log(`[op:writeValues] Sheet "${op.sheet}" not found, creating...`);
          sheet = workbook.worksheets.add(op.sheet);
          sheet.load("name");
          await context.sync();
        }
        const dataRows = op.values.length;
        const dataCols = dataRows > 0 ? op.values[0].length : 0;
        const adjustedRange = adjustRangeToData(sheet, op.range, dataRows, dataCols);
        console.log(
          `[op:writeValues] range "${op.range}" → adjusted to "${adjustedRange}" for ${dataRows}×${dataCols} data`
        );
        const range = sheet.getRange(adjustedRange);
        range.load("address");
        range.values = op.values;
        await context.sync();
        console.log(`[op:writeValues] Wrote ${dataRows} rows to ${op.sheet}!${adjustedRange}`);
        return `Wrote ${dataRows} rows to ${op.sheet}!${adjustedRange}`;
      } catch (e) {
        console.error(`[op:writeValues] Failed:`, e.message);
        throw e;
      }
    }

    case "writeFormulas": {
      console.log(
        `[op:writeFormulas] sheet="${op.sheet}" range="${op.range}" rows=${op.formulas.length}`
      );
      try {
        let sheet;
        try {
          sheet = workbook.worksheets.getItem(op.sheet);
          sheet.load("name");
          await context.sync();
        } catch {
          sheet = workbook.worksheets.add(op.sheet);
          sheet.load("name");
          await context.sync();
        }
        const dataRows = op.formulas.length;
        const dataCols = dataRows > 0 ? op.formulas[0].length : 0;
        const adjustedRange = adjustRangeToData(sheet, op.range, dataRows, dataCols);
        console.log(
          `[op:writeFormulas] range "${op.range}" → adjusted to "${adjustedRange}" for ${dataRows}×${dataCols} data`
        );
        const range = sheet.getRange(adjustedRange);
        range.formulas = op.formulas;
        await context.sync();
        return `Wrote ${dataRows} formulas to ${op.sheet}!${adjustedRange}`;
      } catch (e) {
        console.error(`[op:writeFormulas] Failed:`, e.message);
        throw e;
      }
    }

    case "readRange": {
      console.log(`[op:readRange] sheet="${op.sheet}" range="${op.range}"`);
      let sheet12;
      try {
        sheet12 = workbook.worksheets.getItem(op.sheet);
      } catch {
        const sheets = workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        const availableSheets = sheets.items.map((s) => s.name).join(", ");
        throw new Error(
          `Sheet "${op.sheet}" does not exist. Available sheets: [${availableSheets}]. ` +
            `Please use one of the available sheet names.`
        );
      }
      sheet12.load("name");
      await context.sync();
      const range12 = sheet12.getRange(op.range);
      range12.load("values");
      await context.sync();
      const values = range12.values;
      const normalized = values.map((row) =>
        row.map((cell) => (cell === undefined || cell === null ? "" : cell))
      );
      const rows = normalized.length;
      const cols = normalized[0] ? normalized[0].length : 0;
      const sample = [];
      const displayRows = Math.min(rows, 5);
      for (let r = 0; r < displayRows; r++) {
        const rowCells = [];
        for (let c = 0; c < cols; c++) {
          const val = normalized[r][c];
          rowCells.push(val === "" ? "(empty)" : JSON.stringify(val));
        }
        sample.push(rowCells.join(" | "));
      }
      const actualSheetName = sheet12.name;
      const summary = `${rows} rows x ${cols} cols on ${actualSheetName}!${op.range}: ${sample.join(" | ")}`;
      console.log(`[op:readRange] ${summary}`);
      return summary;
    }

    // === Cell Formatting ===
    case "setNumberFormat": {
      const sheet1 = workbook.worksheets.getItem(op.sheet);
      const range1 = sheet1.getRange(op.range);
      range1.numberFormat = op.format;
      await context.sync();
      return `Set number format "${op.format}" on ${op.sheet}!${op.range}`;
    }

    case "setFillColor": {
      const sheet2 = workbook.worksheets.getItem(op.sheet);
      const range2 = sheet2.getRange(op.range);
      range2.format.fill.color = op.color;
      await context.sync();
      return `Set fill color ${op.color} on ${op.sheet}!${op.range}`;
    }

    case "setFontColor": {
      const sheet3 = workbook.worksheets.getItem(op.sheet);
      const range3 = sheet3.getRange(op.range);
      range3.format.font.color = op.color;
      await context.sync();
      return `Set font color ${op.color} on ${op.sheet}!${op.range}`;
    }

    case "setFontBold": {
      const sheet4 = workbook.worksheets.getItem(op.sheet);
      const range4 = sheet4.getRange(op.range);
      range4.format.font.bold = op.bold;
      await context.sync();
      return `Set font bold=${op.bold} on ${op.sheet}!${op.range}`;
    }

    case "setFontItalic": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      range.format.font.italic = op.italic;
      await context.sync();
      return `Set font italic=${op.italic} on ${op.sheet}!${op.range}`;
    }

    case "setFontUnderline": {
const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      const underlineStyle = Excel.FontUnderlineStyle;
      range.format.font.underline = op.underline
        ? (underlineStyle ? underlineStyle.singleLine : 1)
        : (underlineStyle ? underlineStyle.none : 0);
      await context.sync();
      return `Set font underline=${op.underline} on ${op.sheet}!${op.range}`;
    }

    case "setFontStrikethrough": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      range.format.font.strikethrough = op.strikethrough;
      await context.sync();
      return `Set font strikethrough=${op.strikethrough} on ${op.sheet}!${op.range}`;
    }

    case "setFontName": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      range.format.font.name = op.font;
      await context.sync();
      return `Set font name "${op.font}" on ${op.sheet}!${op.range}`;
    }

    case "setFontSize": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      range.format.font.size = op.size;
      await context.sync();
      return `Set font size ${op.size} on ${op.sheet}!${op.range}`;
    }

    case "setHorizontalAlignment": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      const hAlign = getHAlignMap();
      let hAlignValue = hAlign[op.alignment];
      if (hAlignValue === undefined) {
        try {
          hAlignValue = Excel.HorizontalAlignment ? Excel.HorizontalAlignment.general : 0;
        } catch {
          hAlignValue = 0;
        }
      }
      range.load("format/alignment/horizontal");
      await context.sync();
      range.format.alignment.horizontal = hAlignValue;
      await context.sync();
      return `Set horizontal alignment "${op.alignment}" on ${op.sheet}!${op.range}`;
    }

    case "setVerticalAlignment": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      const vAlign = getVAlignMap();
      let vAlignValue = vAlign[op.alignment];
      if (vAlignValue === undefined) {
        try {
          vAlignValue = Excel.VerticalAlignment ? Excel.VerticalAlignment.bottom : 0;
        } catch {
          vAlignValue = 0;
        }
      }
      range.format.alignment.vertical = vAlignValue;
      await context.sync();
      return `Set vertical alignment "${op.alignment}" on ${op.sheet}!${op.range}`;
    }

    case "setTextWrap": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      range.format.alignment.wrapText = !!op.wrap;
      await context.sync();
      return `Set text wrap=${op.wrap} on ${op.sheet}!${op.range}`;
    }

    case "setIndentLevel": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      range.format.alignment.indentLevel = op.level;
      await context.sync();
      return `Set indent level ${op.level} on ${op.sheet}!${op.range}`;
    }

    case "setRotationAngle": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      range.format.alignment.textRotation = op.angle;
      await context.sync();
      return `Set rotation angle ${op.angle} on ${op.sheet}!${op.range}`;
    }

    case "setBorder": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      const locMap = getBorderLocationMap();
      const styleMap = getBorderStyleMap();
      // Map border location string to index: edgeTop=0, edgeBottom=1, edgeLeft=2, edgeRight=3, insideHorizontal=4, insideVertical=5
      const locationIndex = locMap[op.location] !== undefined ? locMap[op.location] : 0;
      const location = locMap[op.location] || (Excel.BorderLocation ? Excel.BorderLocation.edgeTop : 0);
      const style = styleMap[op.style] || (Excel.BorderStyle ? Excel.BorderStyle.continuous : 0);
      const border = range.format.borders.getItem(locationIndex);
      border.style = style;
      border.color = op.color;
      if (Excel.BorderWeight) {
        border.weight = Excel.BorderWeight.thin;
      }
      await context.sync();
      return `Set border ${op.location} (${op.style}, ${op.color}) on ${op.sheet}!${op.range}`;
    }

    case "setMergeCells": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      if (op.merge) {
        range.merge(true);
      } else {
        range.unmerge();
      }
      await context.sync();
      return `${op.merge ? "Merged" : "Unmerged"} ${op.sheet}!${op.range}`;
    }

    case "setHyperlink": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      range.hyperlink = op.link;
      await context.sync();
      return `Set hyperlink "${op.link}" on ${op.sheet}!${op.range}`;
    }

    case "setCellProtection": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      range.format.protection.locked = op.locked;
      if (typeof op.hidden !== "undefined") {
        range.format.protection.hidden = op.hidden;
      }
      await context.sync();
      return `Set protection locked=${op.locked} on ${op.sheet}!${op.range}`;
    }

    // === Autofit ===
    case "autofitColumns": {
      const sheet5 = workbook.worksheets.getItem(op.sheet);
      sheet5.getUsedRange().format.autofitColumns();
      await context.sync();
      return `Autofit columns on "${op.sheet}"`;
    }

    case "autofitRows": {
      const sheet6 = workbook.worksheets.getItem(op.sheet);
      sheet6.getUsedRange().format.autofitRows();
      await context.sync();
      return `Autofit rows on "${op.sheet}"`;
    }

    // === Tables ===
    case "createTable": {
      const sheet7 = workbook.worksheets.getItem(op.sheet);
      const table = sheet7.tables.add(op.range, true);
      table.name = op.name;
      await context.sync();
      return `Created table "${op.name}" on ${op.sheet}!${op.range}`;
    }

    case "deleteTable": {
      const table = workbook.tables.getItemOrNullObject(op.name);
      table.load("name, isNullObject");
      await context.sync();
      if (table.isNullObject) {
        return `Skipped delete: table "${op.name}" not found`;
      }
      table.delete();
      await context.sync();
      return `Deleted table "${op.name}"`;
    }

    case "addTableRow": {
      const table = workbook.tables.getItem(op.tableName);
      table.rows.add(op.index, op.values);
      await context.sync();
      return `Added row to table "${op.tableName}" at index ${op.index}`;
    }

    case "addTableColumn": {
      const table = workbook.tables.getItem(op.tableName);
      table.columns.add(op.name);
      await context.sync();
      return `Added column "${op.name}" to table "${op.tableName}"`;
    }

    case "deleteTableRow": {
      const table = workbook.tables.getItem(op.tableName);
      const row = table.rows.getItemAt(op.index);
      row.delete();
      await context.sync();
      return `Deleted row at index ${op.index} from table "${op.tableName}"`;
    }

    case "deleteTableColumn": {
      const table = workbook.tables.getItem(op.tableName);
      const col = table.columns.getItemAt(op.index);
      col.delete();
      await context.sync();
      return `Deleted column at index ${op.index} from table "${op.tableName}"`;
    }

    case "sortTable": {
      const table = workbook.tables.getItem(op.tableName);
      const keys = op.columns.map((c) => {
        const sortOrder = getSortOrderMap();
        return {
          key: c.key,
          ascending: c.order ? (sortOrder[c.order] || 0) : (Excel.SortOrder ? Excel.SortOrder.ascending : 0),
        };
      });
      table.sort.apply(keys);
      await context.sync();
      return `Sorted table "${op.tableName}" by keys: ${keys.map((k) => `key=${k.key} ${k.ascending ? "↑" : "↓"}`).join(", ")}`;
    }

    case "filterTable": {
      const table = workbook.tables.getItem(op.tableName);
      const columnIndex = op.columnIndex;
      let filter;
      if (op.rule && op.rule.type === "containsText") {
        filter = {
          field: columnIndex,
          calculatedField: false,
          filterType: Excel.FilterType.dynamic,
          dynamicType: Excel.DynamicFilterType.containsText,
          criterion: op.rule.value,
        };
      } else if (op.rule && op.rule.type === "valueFilter") {
        filter = {
          field: columnIndex,
          calculatedField: false,
          filterType: Excel.FilterType.values,
          values: op.rule.values || [],
        };
      } else if (op.rule && op.rule.type === "topBottom") {
        filter = {
          field: columnIndex,
          calculatedField: false,
          filterType: Excel.FilterType.topBottom,
          topBottomType: op.rule.above ? Excel.TopBottom.above : Excel.TopBottom.below,
          count: op.rule.count || 10,
        };
      } else {
        filter = {
          field: columnIndex,
          calculatedField: false,
          filterType: Excel.FilterType.dynamic,
          dynamicType: Excel.DynamicFilterType.none,
        };
      }
      table.filters.add(filter);
      await context.sync();
      return `Added filter on table "${op.tableName}" column ${columnIndex}`;
    }

    case "setTableStyle": {
      const table = workbook.tables.getItem(op.tableName);
      table.style = op.style;
      await context.sync();
      return `Applied style "${op.style}" to table "${op.tableName}"`;
    }

    case "setTableShowHeader": {
      const table = workbook.tables.getItem(op.tableName);
      table.showHeader = op.show;
      await context.sync();
      return `Set header visible=${op.show} on table "${op.tableName}"`;
    }

    case "setTableShowTotals": {
      const table = workbook.tables.getItem(op.tableName);
      table.showTotals = op.show;
      await context.sync();
      return `Set totals visible=${op.show} on table "${op.tableName}"`;
    }

    case "setTableBandedRows": {
      const table = workbook.tables.getItem(op.tableName);
      table.bandedRows = op.banded;
      await context.sync();
      return `Set banded rows=${op.banded} on table "${op.tableName}"`;
    }

    case "setTableBandedColumns": {
      const table = workbook.tables.getItem(op.tableName);
      table.bandedColumns = op.banded;
      await context.sync();
      return `Set banded columns=${op.banded} on table "${op.tableName}"`;
    }

    // === Charts ===
    case "addChart": {
      const sheet8 = workbook.worksheets.getItem(op.sheet);
      const chartTypeMap = getChartTypeMap();
      const chartType = chartTypeMap[op.chartType] || (Excel.ChartType ? Excel.ChartType.columnClustered : 0);
      const chartSeriesBy = Excel.ChartSeriesBy;
      const chart = sheet8.charts.add(chartType, sheet8.getRange(op.dataRange), chartSeriesBy ? chartSeriesBy.columns : 0);
      chart.name = op.title;
      chart.title.text = op.title;
      chart.title.format.font.size = 12;
      chart.left = op.left || 200;
      chart.top = op.top || 100;
      chart.width = op.width || 300;
      chart.height = op.height || 200;
      if (typeof op.showLegend !== "undefined") {
        chart.legend.visible = op.showLegend;
      }
      await context.sync();
      return `Created chart "${op.title}" on ${op.sheet}`;
    }

    case "deleteChart": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const chart = sheet.charts.getItemOrNullObject(op.name);
      chart.load("name, isNullObject");
      await context.sync();
      if (chart.isNullObject) {
        return `Skipped delete: chart "${op.name}" not found on ${op.sheet}`;
      }
      chart.delete();
      await context.sync();
      return `Deleted chart "${op.name}" from ${op.sheet}`;
    }

    case "setChartTitle": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const chart = sheet.charts.getItem(op.name);
      chart.title.text = op.title;
      await context.sync();
      return `Set chart title to "${op.title}" on ${op.sheet}`;
    }

    case "setChartSeries": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const chart = sheet.charts.getItem(op.name);
      chart.serieses.getItem(0).values = sheet.getRange(op.dataRange);
      if (op.seriesName) {
        chart.serieses.getItem(0).name = op.seriesName;
      }
      await context.sync();
      return `Updated chart series on ${op.sheet}`;
    }

    case "setChartPosition": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const chart = sheet.charts.getItem(op.name);
      chart.left = op.left || chart.left;
      chart.top = op.top || chart.top;
      chart.width = op.width || chart.width;
      chart.height = op.height || chart.height;
      await context.sync();
      return `Set chart position on ${op.sheet}`;
    }

    case "setChartShowLegend": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const chart = sheet.charts.getItem(op.name);
      chart.legend.visible = op.show;
      await context.sync();
      return `Set legend visible=${op.show} on chart "${op.name}"`;
    }

    case "setChartShowDataLabels": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const chart = sheet.charts.getItem(op.name);
      chart.serieses.getItem(0).hasDataLabels = op.show;
      await context.sync();
      return `Set data labels visible=${op.show} on chart "${op.name}"`;
    }

    // === PivotTables ===
    case "createPivotTable": {
      const pivotSheet = workbook.worksheets.getItem(op.pivotSheet);
      // sourceData is a string like "Sheet1!$A$1:$D$100", not a Range object
      const pivotTable = pivotSheet.pivotTables.add(op.name, op.sourceData, op.pivotRange);
      console.log(`[op:createPivotTable] Created "${op.name}"`);
      return `Created pivot table "${op.name}" on ${op.pivotSheet}!${op.pivotRange}`;
    }

    case "addPivotHierarchy": {
      const pivotTable = workbook.pivotTables.getItem(op.pivotTable);
      // Load the hierarchies we need
      pivotTable.rowHierarchies.load("items/name");
      pivotTable.columnHierarchies.load("items/name");
      pivotTable.dataHierarchies.load("items/name");
      pivotTable.filterHierarchies.load("items/name");
      await context.sync();

      let hierarchy;
      let fieldIndex = -1;
      switch (op.category) {
        case "row":
          hierarchy = pivotTable.rowHierarchies;
          fieldIndex = hierarchy.items.findIndex((f) => f.name === op.field);
          break;
        case "column":
          hierarchy = pivotTable.columnHierarchies;
          fieldIndex = hierarchy.items.findIndex((f) => f.name === op.field);
          break;
        case "data":
          hierarchy = pivotTable.dataHierarchies;
          fieldIndex = hierarchy.items.findIndex((f) => f.name === op.field);
          break;
        case "filter":
          hierarchy = pivotTable.filterHierarchies;
          fieldIndex = hierarchy.items.findIndex((f) => f.name === op.field);
          break;
        default:
          throw new Error(`Unknown pivot category: "${op.category}"`);
      }
      if (fieldIndex < 0) {
        throw new Error(`Field "${op.field}" not found in pivot table "${op.pivotTable}"`);
      }
      const insertPosEnum = getPivotHierarchyInsertPosition();
      const insertPos = op.index !== undefined ? op.index : (insertPosEnum ? insertPosEnum.end : 1);
      hierarchy.add(insertPos, hierarchy.items[fieldIndex]);
      await context.sync();
      return `Added field "${op.field}" to "${op.category}" of pivot table "${op.pivotTable}"`;
    }

    case "setPivotAggregation": {
      const pivotTable = workbook.pivotTables.getItem(op.pivotTable);
      // Load the hierarchies we need
      pivotTable.rowHierarchies.load("items/name");
      pivotTable.columnHierarchies.load("items/name");
      pivotTable.dataHierarchies.load("items/name");
      pivotTable.filterHierarchies.load("items/name");
      await context.sync();

      let hierarchy;
      let fieldIndex = -1;
      switch (op.category) {
        case "row":
          hierarchy = pivotTable.rowHierarchies;
          fieldIndex = hierarchy.items.findIndex((f) => f.name === op.field);
          break;
        case "column":
          hierarchy = pivotTable.columnHierarchies;
          fieldIndex = hierarchy.items.findIndex((f) => f.name === op.field);
          break;
        case "data":
          hierarchy = pivotTable.dataHierarchies;
          fieldIndex = hierarchy.items.findIndex((f) => f.name === op.field);
          break;
        case "filter":
          hierarchy = pivotTable.filterHierarchies;
          fieldIndex = hierarchy.items.findIndex((f) => f.name === op.field);
          break;
        default:
          throw new Error(`Unknown pivot category: "${op.category}"`);
      }
      if (fieldIndex < 0) {
        throw new Error(`Field "${op.field}" not found in pivot table "${op.pivotTable}"`);
      }
      const aggMap = getPivotAggregationMap();
      field.dataAggregation = aggMap
        ? aggMap[op.aggregation] || aggMap.sum
        : Excel.PivotDataAggregation
          ? Excel.PivotDataAggregation.sum
          : 0;
      await context.sync();
      return `Set aggregation "${op.aggregation}" for field "${op.field}" in "${op.pivotTable}"`;
    }

    case "setPivotLayout": {
      const pivotTable = workbook.pivotTables.getItem(op.pivotTable);
      const layoutMap = getPivotLayoutMap();
      pivotTable.layout = layoutMap
        ? layoutMap[op.layout] || layoutMap.tabular
        : Excel.PivotLayoutType
          ? Excel.PivotLayoutType.tabular
          : 0;
      await context.sync();
      return `Set layout "${op.layout}" on pivot table "${op.pivotTable}"`;
    }

    case "addPivotFilter": {
      const pivotTable = workbook.pivotTables.getItem(op.pivotTable);
      const field = pivotTable.getAllFields().find((f) => f.name === op.field);
      if (!field) {
        throw new Error(`Field "${op.field}" not found in pivot table "${op.pivotTable}"`);
      }
      pivotTable.filterHierarchies.add(Excel.PivotHierarchyInsertionPosition.end, field);
      await context.sync();
      return `Added field "${op.field}" to filter area of pivot table "${op.pivotTable}"`;
    }

    case "deletePivotTable": {
      const pt = workbook.pivotTables.getItemOrNullObject(op.name);
      pt.load("name, isNullObject");
      await context.sync();
      if (pt.isNullObject) {
        return `Skipped delete: pivot table "${op.name}" not found`;
      }
      pt.delete();
      await context.sync();
      return `Deleted pivot table "${op.name}"`;
    }

    // === Conditional Formatting ===
    case "addDataBar": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
      const dataBar = cf.dataBar;
      dataBar.color = op.color;
      await context.sync();
      return `Added data bar to ${op.sheet}!${op.range}`;
    }

    case "addColorScale": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
      const colorScale = cf.colorScale;
      colorScale.load("criteria");
      await context.sync();
      if (op.colors && op.colors.length >= 2) {
        colorScale.criteria.add(0, op.colors[0]);
        colorScale.criteria.add(50, op.colors[1]);
        if (op.colors.length >= 3) {
          colorScale.criteria.add(100, op.colors[2]);
        }
      }
      await context.sync();
      return `Added color scale to ${op.sheet}!${op.range}`;
    }

    case "addIconSet": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
      const iconSet = cf.iconSet;
      const iconSets = getIconSetsMap();
      iconSet.iconSet = iconSets[op.iconSet] || (Excel.IconSetType ? Excel.IconSetType.threeArrows : 0);
      await context.sync();
      return `Added icon set "${op.iconSet}" to ${op.sheet}!${op.range}`;
    }

    case "addTextComparison": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.textComparison);
const textCf = cf.textComparison;
      const textMap = getTextComparisonMap();
      const operator = textMap[op.operator] || (Excel.TextComparisonOperator ? Excel.TextComparisonOperator.containsText : 0);
      textCf.criterion = op.text || "";
      textCf.textComparison = textMap[op.textComparison] || (Excel.TextComparisonOperator ? Excel.TextComparisonOperator.containsText : 0);
      textCf.operator = operator;
      await context.sync();
      return `Added text comparison to ${op.sheet}!${op.range}`;
    }

    case "addTopBottom": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
      const topBottom = cf.topBottom;
      topBottom.topBottom = op.topBottom;
      topBottom.value = op.value || 10;
      topBottom.belowAverage = op.belowAverage || false;
      await context.sync();
      return `Added top/bottom rule to ${op.sheet}!${op.range}`;
    }

    case "addPresetCriteria": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);
      const preset = cf.presetCriteria;
      const presetCriteria = getPresetCriteriaMap();
      preset.criteria = presetCriteria[op.criteria] || (Excel.PresetCriteria ? Excel.PresetCriteria.aboveAverage : 0);
      await context.sync();
      return `Added preset criteria "${op.criteria}" to ${op.sheet}!${op.range}`;
    }

    case "addCustomConditionalFormat": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      cf.formula = op.formula;
      await context.sync();
      return `Added custom conditional format to ${op.sheet}!${op.range}`;
    }

    case "clearConditionalFormats": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      range.conditionalFormats.clearAll();
      await context.sync();
      return `Cleared all conditional formats on ${op.sheet}!${op.range}`;
    }

    // === Shapes ===
    case "addShape": {
      const sheet = workbook.worksheets.getItem(op.sheet);
const geoShape = Excel.GeometricShapeType;
      const shapeType =
        op.shapeType === "rectangle" ? (geoShape ? geoShape.rectangle : 0)
        : op.shapeType === "ellipse" ? (geoShape ? geoShape.ellipse : 0)
        : op.shapeType === "triangle" ? (geoShape ? geoShape.rightTriangle : 0)
        : op.shapeType === "star" ? (geoShape ? geoShape.star5 : 0)
        : op.shapeType === "circle" ? (geoShape ? geoShape.ellipse : 0)
        : (geoShape ? geoShape.rectangle : 0);
      const shape = sheet.shapes.addGeometricShape(shapeType, op.left, op.top, op.width, op.height);
      shape.name = op.name;
      await context.sync();
      return `Added shape "${op.name}" to ${op.sheet}`;
    }

    case "addTextBox": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const shape = sheet.shapes.addTextBox(op.text || "", op.left, op.top, op.width, op.height);
      shape.name = op.name;
      await context.sync();
      return `Added text box "${op.name}" to ${op.sheet}`;
    }

    case "setShapeText": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const shape = sheet.shapes.getItem(op.name);
      shape.textFrame.content = op.text;
      await context.sync();
      return `Set text on shape "${op.name}" on ${op.sheet}`;
    }

    case "deleteShape": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const shape = sheet.shapes.getItemOrNullObject(op.name);
      shape.load("name, isNullObject");
      await context.sync();
      if (shape.isNullObject) {
        return `Skipped delete: shape "${op.name}" not found on ${op.sheet}`;
      }
      shape.delete();
      await context.sync();
      return `Deleted shape "${op.name}" from ${op.sheet}`;
    }

    // === Data Validation ===
    case "addDataValidation": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      const ruleMap = getValidationRuleMap();
      const opMap = getValidationOperatorMap();
      const rule = ruleMap[op.ruleType] || (Excel.DataValidationType ? Excel.DataValidationType.wholeNumber : 0);
      const validation = range.dataValidation;
      const opEnum = opMap[op.operator] || (Excel.DataValidationOperator ? Excel.DataValidationOperator.greaterThan : 0);

      // Use dataValidation.set() instead of apply()
      if (op.ruleType === "list") {
        validation.set({
          type: rule,
          ignoreBlanks: true,
          formula1: op.value1 ? op.value1.split(",") : [],
        });
      } else {
        validation.set({
          type: rule,
          operator: opEnum,
          formula1: op.value1 || undefined,
          formula2: op.value2 || undefined,
          ignoreBlanks: true,
        });
      }
      await context.sync();
      return `Added data validation (${op.ruleType}) to ${op.sheet}!${op.range}`;
    }

    case "deleteDataValidation": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      range.dataValidation.clear();
      await context.sync();
      return `Cleared data validation on ${op.sheet}!${op.range}`;
    }

    // === Comments & Notes ===
    case "addComment": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      // CommentCollection.add(cellAddress, content, contentType) - cellAddress is a string like "A1"
      const cellAddress = op.range.split(":")[0] || "A1";
      sheet.comments.add(cellAddress, op.text || "", "Short");
      await context.sync();
      return `Added comment to ${op.sheet}!${cellAddress}`;
    }

    case "deleteComment": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const cellAddress = op.range.split(":")[0] || "A1";
      const comment = sheet.comments.getItemOrNullObject(cellAddress);
      comment.load("isNullObject");
      await context.sync();
      if (!comment.isNullObject) {
        comment.delete();
        await context.sync();
      }
      return `Deleted comment from ${op.sheet}!${cellAddress}`;
    }

    case "addNote": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      // NoteCollection.add(cellAddress, content) - cellAddress is a string like "A1"
      const cellAddress = op.range.split(":")[0] || "A1";
      sheet.notes.add(cellAddress, op.text || "");
      await context.sync();
      return `Added note to ${op.sheet}!${cellAddress}`;
    }

    // === Named Items ===
    case "addNamedRange": {
      workbook.names.add(op.name, op.definition);
      await context.sync();
      return `Added named range "${op.name}" = "${op.definition}"`;
    }

    case "deleteNamedRange": {
      const namedItem = workbook.names.getItemOrNullObject(op.name);
      namedItem.load("isNullObject");
      await context.sync();
      if (!namedItem.isNullObject) {
        namedItem.delete();
        await context.sync();
      }
      return `Deleted named range "${op.name}"`;
    }

    // === Worksheet View & Layout ===
    case "setFreezePanes": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const freezePanes = sheet.freezePanes;
      if (op.freezeType === "freezeRows") {
        freezePanes.freezeRows(op.row || 1);
      } else if (op.freezeType === "freezeColumns") {
        freezePanes.freezeColumns(op.column || 1);
      } else if (op.freezeType === "freezePanes") {
        // freezeAt takes a range or string address
        const frozenRange = op.row && op.column
          ? `${indexToCol(op.column)}${op.row}:${indexToCol(op.column)}${op.row}`
          : `A1`;
        freezePanes.freezeAt(frozenRange);
      } else if (op.freezeType === "unfreeze") {
        freezePanes.unfreeze();
      }
      await context.sync();
      return `Set freeze panes (${op.freezeType}) on "${op.sheet}"`;
    }

    case "setPageLayout": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const pageLayout = sheet.getPageLayout();
      if (op.orientation) {
        pageLayout.orientation =
          op.orientation === "landscape"
            ? Excel.PageOrientation.landscape
            : Excel.PageOrientation.portrait;
      }
      if (op.pageSize) {
        pageLayout.page_size = op.pageSize === "a4" ? Excel.PageSize.A4 : Excel.PageSize.letter;
      }
      await context.sync();
      return `Set page layout on "${op.sheet}"`;
    }

    case "setPrintSettings": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      if (op.printArea) {
        sheet.setPrintArea(op.printArea);
      }
      await context.sync();
      return `Set print settings on "${op.sheet}"`;
    }

    case "addPageBreak": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      if (op.position) {
        const match = op.position.match(/^([A-Z]+)(\d+)/i);
        if (match) {
          const col = match[1].toUpperCase();
          const row = parseInt(match[2], 10);
          sheet.horizontalPageBreaks.add(row);
        }
      }
      await context.sync();
      return `Added page break on "${op.sheet}"`;
    }

    // === Protection ===
    case "protectWorksheet": {
      const sheet = workbook.worksheets.getItem(op.name);
      sheet.protect(op.password || "");
      await context.sync();
      return `Protected sheet "${op.name}"`;
    }

    case "unprotectWorksheet": {
      const sheet = workbook.worksheets.getItem(op.name);
      sheet.unprotect(op.password || "");
      await context.sync();
      return `Unprotected sheet "${op.name}"`;
    }

// === Sort ===
    case "sortRange": {
      const sheet = workbook.worksheets.getItem(op.sheet);
      const range = sheet.getRange(op.range);
      const sortOrder = getSortOrderMap();
      const keys = [
        {
          key: op.key,
          ascending: op.order ? (sortOrder[op.order] || 0) : (Excel.SortOrder ? Excel.SortOrder.ascending : 0),
        },
      ];
      range.sort.apply(keys);
      await context.sync();
      return `Sorted range ${op.sheet}!${op.range} by key ${op.key}`;
    }

    // === Delete ===
    case "deleteRows": {
      console.log(`[op:deleteRows] sheet="${op.sheet}" range="${op.range}"`);
      const sheet10 = workbook.worksheets.getItem(op.sheet);
      sheet10.load("name");
      await context.sync();
      const range10 = sheet10.getRange(op.range);
      const deleteShift = Excel.DeleteShiftDirection;
      range10.delete(deleteShift ? deleteShift.up : 0);
      await context.sync();
      return `Deleted rows in ${op.sheet}!${op.range}`;
    }

    case "deleteColumns": {
      console.log(`[op:deleteColumns] sheet="${op.sheet}" range="${op.range}"`);
      const sheet11 = workbook.worksheets.getItem(op.sheet);
      sheet11.load("name");
      await context.sync();
      const range11 = sheet11.getRange(op.range);
      const deleteShift = Excel.DeleteShiftDirection;
      range11.delete(deleteShift ? deleteShift.left : 1);
      await context.sync();
      return `Deleted columns in ${op.sheet}!${op.range}`;
    }

    // === Slicers ===
    case "addSlicer": {
      const pivotTable = workbook.pivotTables.getItem(op.pivotTable);
      const field = pivotTable.getAllFields().find((f) => f.name === op.field);
      if (!field) {
        throw new Error(`Field "${op.field}" not found in pivot table "${op.pivotTable}"`);
      }
      const slicerSheet = workbook.worksheets.getItem(op.sheet);
      const slicer = slicerSheet.slicers.add(op.slicerName, pivotTable, field);
      await context.sync();
      return `Added slicer "${op.slicerName}" for field "${op.field}"`;
    }

    case "deleteSlicer": {
       const sheet = workbook.worksheets.getItem(op.sheet);
       const slicer = sheet.slicers.getItemOrNullObject(op.name);
       slicer.load("name, isNullObject");
       await context.sync();
       if (slicer.isNullObject) {
         return `Skipped delete: slicer "${op.name}" not found on ${op.sheet}`;
       }
       slicer.delete();
       await context.sync();
       return `Deleted slicer "${op.name}" from ${op.sheet}`;
     }

     case "clearSheet": {
       const sheet = workbook.worksheets.getItem(op.sheet);
       // Clear all used ranges
       const usedRange = sheet.getUsedRange();
       if (!usedRange.isNullObject) {
         usedRange.clear(Excel.ClearApplyTo.all);
       }
       // Delete all charts
       const charts = sheet.charts;
       charts.load("items/name");
       await context.sync();
       for (const chart of charts.items) {
         chart.delete();
       }
       // Delete all shapes
       const shapes = sheet.shapes;
       shapes.load("items/name");
       await context.sync();
       for (const shape of shapes.items) {
         shape.delete();
       }
       // Delete all tables
       const tables = sheet.tables;
       tables.load("items/name");
       await context.sync();
       for (const table of tables.items) {
         table.delete();
       }
       // Clear conditional formatting
       sheet.clearConditionalFormats();
       // Clear filters
       sheet.autoFilter.clear();
       await context.sync();
       return `Cleared all data, formatting, charts, shapes, tables, and filters from "${op.sheet}"`;
     }

     default:
      throw new Error(`Unknown operation type: "${type}"`);
  }
}

/* Dispatcher — routes executeOperation to the correct host engine */
let _pptExecuteOperation = null;

export function setPptExecuteOperation(fn) {
  _pptExecuteOperation = fn;
}

// ─── Retry Wrapper ────────────────────────────────────────────────────────────

/**
 * Retry a failed operation once with exponential backoff.
 * Catches transient errors (timeout, network) and retries.
 */
async function executeWithRetry(executeFn, opName, maxRetries = 1) {
  let lastError;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const result = await executeFn();
      return result;
    } catch (error) {
      lastError = error;

      // Check if this is a transient error worth retrying
      const isTransient =
        error.message?.includes("timeout") ||
        error.message?.includes("network") ||
        error.message?.includes("sync") ||
        error.message?.includes("429") ||
        error.message?.includes("rate limit");

      if (attempt < maxRetries && isTransient) {
        const backoffMs = 1000 * (attempt + 1); // 1s, 2s, ...
        console.warn(
          `[op:${opName}] Transient error (attempt ${attempt + 1}/${maxRetries + 1}): ${error.message}. ` +
          `Retrying in ${backoffMs}ms...`
        );
        await new Promise((resolve) => setTimeout(resolve, backoffMs));
        continue;
      }

      // Non-transient or max retries reached — throw
      if (attempt === maxRetries) {
        console.error(`[op:${opName}] Failed after ${maxRetries + 1} attempts: ${error.message}`);
      }
      throw error;
    }
  }

  // Should not reach here, but just in case
  throw lastError;
}

// ─── Public API ──────────────────────────────────────────────────────────────

export function getHostRegistry(host) {
  if (host === "PowerPoint") {
    return null; // PPT registry is in ppt/operations.js
  }
  return EXCEL_OPERATION_REGISTRY;
}

export async function dispatchExecuteOperation(host, context, op) {
  if (host === "PowerPoint" && _pptExecuteOperation) {
    return _pptExecuteOperation(context, op);
  }
  return executeOperation(context, op);
}

/**
 * Execute an operation with retry logic and telemetry tracking.
 * This is the public entry point — wraps executeOperation with retry + error tracking.
 */
export async function executeOperationWithTracking(context, op) {
  const opName = op.operation;

  // Execute with retry
  const result = await executeWithRetry(
    async () => executeOperation(context, op),
    opName,
    1 // 1 retry max
  );

  // Record success in telemetry (via global telemetry module)
  try {
    const telemetry = await import("../telemetry.js");
    telemetry.recordOpResult(opName, true);
  } catch {
    // Telemetry not available — ignore
  }

  return result;
}
