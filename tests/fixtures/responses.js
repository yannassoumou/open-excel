/**
 * Sample LLM responses for parser tests.
 * Both valid and invalid formats.
 */

// Valid: standard JSON code block
export const validStandardBlock = `\`\`\`json
{
  "plan": "Create a summary sheet and write data",
  "operations": [
    {"operation": "createWorksheet", "name": "Summary"},
    {"operation": "writeValues", "sheet": "Summary", "range": "A1:C3", "values": [["Name", "Value", "Note"], ["Item1", 42, "ok"]]}
  ]
}
\`\`\``;

// Valid: JSON without language tag
export const validGenericBlock = `\`\`\`
{
  "plan": "Write some data",
  "operations": [
    {"operation": "writeValues", "sheet": "Sheet1", "range": "A1:A1", "values": [["test"]]}
  ]
}
\`\`\``;

// Valid: no JSON block (plain text)
export const validPlainText = `I'll help you with that. Let me start by creating a worksheet.`;

// Valid: JSON with extra text outside
export const validExtraText = `Here's the plan:\n\n\`\`\`json
{
  "plan": "Create sheet",
  "operations": [{"operation": "createWorksheet", "name": "Data"}]
}
\`\`\`\n\nLet me know if you need more.`;

// Valid: JSON without newline after ```
export const validNoNewline = `\`\`\`json{"plan": "test", "operations": []}`;

// Invalid: malformed JSON
export const invalidMalformedJSON = `\`\`\`json
{
  "plan": "test",
  "operations": [{"operation": "createWorksheet", "name": "Data"}
}
\`\`\``;

// Invalid: no plan or operations field
export const invalidMissingFields = `\`\`\`json
{
  "message": "I'll do it"
}
\`\`\``;

// Invalid: operations is not an array
export const invalidOpsNotArray = `\`\`\`json
{
  "plan": "test",
  "operations": {"operation": "createWorksheet", "name": "Data"}
}
\`\`\``;

// Invalid: empty response
export const invalidEmpty = ``;

// Invalid: null response
export const invalidNull = null;

// Invalid: operations array is empty
export const validEmptyOps = `\`\`\`json
{
  "plan": "All done",
  "operations": []
}
\`\`\``;

// Invalid: unknown operation type
export const invalidUnknownOp = `\`\`\`json
{
  "plan": "Do something",
  "operations": [{"operation": "magicSpell", "name": "Abracadabra"}]
}
\`\`\``;

// Invalid: missing required field
export const invalidMissingField = `\`\`\`json
{
  "plan": "Write data",
  "operations": [{"operation": "writeValues", "sheet": "Sheet1"}]
}
\`\`\``;

// Invalid: wrong field type
export const invalidWrongType = `\`\`\`json
{
  "plan": "Bold text",
  "operations": [{"operation": "setFontBold", "sheet": "Sheet1", "range": "A1", "bold": "yes"}]
}
\`\`\``;

// Valid: multiple operations in one block
export const validMultiOps = `\`\`\`json
{
  "plan": "Create sheet, write data, format, and add chart",
  "operations": [
    {"operation": "createWorksheet", "name": "Dashboard"},
    {"operation": "writeValues", "sheet": "Dashboard", "range": "A1:B1", "values": [["Metric", "Value"]]},
    {"operation": "setFillColor", "sheet": "Dashboard", "range": "A1:B1", "color": "#4472C4"},
    {"operation": "setFontBold", "sheet": "Dashboard", "range": "A1:B1", "bold": true},
    {"operation": "autofitColumns", "sheet": "Dashboard"}
  ]
}
\`\`\``;

// Valid: AI response with completion claim
export const validCompletionClaim = `\`\`\`json
{
  "plan": "All steps complete",
  "operations": []
}
\`\`\`\n\nAll steps complete.`;

// Valid: AI response with completion claim (no JSON)
export const validCompletionText = `I've finished all the steps. The workbook is ready.`;

// Valid: AI response with _switchMode
export const validSwitchMode = `\`\`\`json
{
  "plan": "Switching to explain mode",
  "operations": [{"operation": "readRange", "sheet": "Sheet1", "range": "A1:C5"}],
  "_switchMode": "explain",
  "_switchMessage": "Now in explain mode"
}
\`\`\``;

// Edge case: JSON with comments (not valid JSON but some LLMs do this)
export const invalidJSONWithComments = `\`\`\`json
{
  // This is a comment
  "plan": "test",
  "operations": []
}
\`\`\``;

// Edge case: operations with null values
export const invalidNullValues = `\`\`\`json
{
  "plan": "test",
  "operations": [
    {"operation": "createWorksheet", "name": null},
    {"operation": "writeValues", "sheet": "Sheet1", "range": "A1", "values": null}
  ]
}
\`\`\``;

// Edge case: very long response with JSON embedded
export const validLongResponse = `Let me think about this step by step.\n\nFirst, I need to understand what the user wants.\n\nHere's my plan:\n\n\`\`\`json
{
  "plan": "Create a comprehensive dashboard",
  "operations": [
    {"operation": "createWorksheet", "name": "Data"},
    {"operation": "createWorksheet", "name": "Dashboard"},
    {"operation": "writeValues", "sheet": "Data", "range": "A1:C1", "values": [["Name", "Value", "Date"]]}
  ]
}
\`\`\`\n\nNow let me execute these operations.`;

// Edge case: multiple JSON blocks (should extract first valid one)
export const invalidMultipleBlocks = `\`\`\`text
not json
\`\`\`\n\n\`\`\`json
{
  "plan": "test",
  "operations": [{"operation": "createWorksheet", "name": "Data"}]
}
\`\`\``;

// Valid: setBorder with all required fields
export const validSetBorder = `\`\`\`json
{
  "plan": "Add borders",
  "operations": [
    {"operation": "setBorder", "sheet": "Sheet1", "range": "A1:C3", "location": "edgeTop", "style": "continuous", "color": "#000000"}
  ]
}
\`\`\``;

// Valid: addChart with all required fields
export const validAddChart = `\`\`\`json
{
  "plan": "Add chart",
  "operations": [
    {"operation": "addChart", "sheet": "Sheet1", "dataRange": "A1:B5", "chartType": "column", "title": "Sales"}
  ]
}
\`\`\``;

// Valid: createPivotTable with all required fields
export const validCreatePivot = `\`\`\`json
{
  "plan": "Create pivot",
  "operations": [
    {"operation": "createPivotTable", "name": "Pivot1", "sourceData": "Sheet1!A1:D100", "pivotSheet": "PivotSheet", "pivotRange": "A1"}
  ]
}
\`\`\``;

// Valid: setNumberFormat
export const validSetNumberFormat = `\`\`\`json
{
  "plan": "Format numbers",
  "operations": [
    {"operation": "setNumberFormat", "sheet": "Sheet1", "range": "B2:B10", "format": "#,##0.00"}
  ]
}
\`\`\``;

// Valid: setHorizontalAlignment
export const validSetHAlign = `\`\`\`json
{
  "plan": "Align text",
  "operations": [
    {"operation": "setHorizontalAlignment", "sheet": "Sheet1", "range": "A1:C1", "alignment": "center"}
  ]
}
\`\`\``;

// Valid: setMergeCells
export const validSetMerge = `\`\`\`json
{
  "plan": "Merge cells",
  "operations": [
    {"operation": "setMergeCells", "sheet": "Sheet1", "range": "A1:C1", "merge": true}
  ]
}
\`\`\``;

// Valid: setFontSize
export const validSetFontSize = `\`\`\`json
{
  "plan": "Set font size",
  "operations": [
    {"operation": "setFontSize", "sheet": "Sheet1", "range": "A1", "size": 14}
  ]
}
\`\`\``;

// Valid: setFontName
export const validSetFontName = `\`\`\`json
{
  "plan": "Set font",
  "operations": [
    {"operation": "setFontName", "sheet": "Sheet1", "range": "A1", "font": "Arial"}
  ]
}
\`\`\``;

// Valid: addDataBar
export const validAddDataBar = `\`\`\`json
{
  "plan": "Add data bar",
  "operations": [
    {"operation": "addDataBar", "sheet": "Sheet1", "range": "B2:B10", "color": "#4472C4", "minType": "min", "minValue": 0, "maxType": "max", "maxValue": 100}
  ]
}
\`\`\``;

// Valid: addComment
export const validAddComment = `\`\`\`json
{
  "plan": "Add comment",
  "operations": [
    {"operation": "addComment", "sheet": "Sheet1", "range": "A1", "text": "Review this data"}
  ]
}
\`\`\``;

// Valid: addShape
export const validAddShape = `\`\`\`json
{
  "plan": "Add shape",
  "operations": [
    {"operation": "addShape", "sheet": "Sheet1", "shapeType": "rectangle", "name": "Box1", "left": 100, "top": 100, "width": 200, "height": 100}
  ]
}
\`\`\``;

// Valid: addDataValidation
export const validAddValidation = `\`\`\`json
{
  "plan": "Add validation",
  "operations": [
    {"operation": "addDataValidation", "sheet": "Sheet1", "range": "C2:C100", "ruleType": "wholeNumber", "operator": "greaterThan", "value1": 0}
  ]
}
\`\`\``;

// Valid: renameWorksheet
export const validRenameSheet = `\`\`\`json
{
  "plan": "Rename sheet",
  "operations": [
    {"operation": "renameWorksheet", "name": "Sheet1", "newName": "Sales Data"}
  ]
}
\`\`\``;

// Valid: hideWorksheet
export const validHideSheet = `\`\`\`json
{
  "plan": "Hide sheet",
  "operations": [
    {"operation": "hideWorksheet", "name": "Sheet2"}
  ]
}
\`\`\``;

// Valid: copyWorksheet
export const validCopySheet = `\`\`\`json
{
  "plan": "Copy sheet",
  "operations": [
    {"operation": "copyWorksheet", "name": "Sheet1", "targetSheet": "Sheet1_Copy", "after": "Sheet1"}
  ]
}
\`\`\``;

// Valid: sortRange
export const validSortRange = `\`\`\`json
{
  "plan": "Sort data",
  "operations": [
    {"operation": "sortRange", "sheet": "Sheet1", "range": "A1:C100", "key": 1, "order": "ascending"}
  ]
}
\`\`\``;

// Valid: createTable
export const validCreateTable = `\`\`\`json
{
  "plan": "Create table",
  "operations": [
    {"operation": "createTable", "sheet": "Sheet1", "range": "A1:D100", "name": "SalesTable"}
  ]
}
\`\`\``;

// Valid: addTableRow
export const validAddTableRow = `\`\`\`json
{
  "plan": "Add row",
  "operations": [
    {"operation": "addTableRow", "tableName": "SalesTable", "index": -1, "values": ["New Item", 100, "Q1", "Active"]}
  ]
}
\`\`\``;

// Valid: deleteRows
export const validDeleteRows = `\`\`\`json
{
  "plan": "Delete rows",
  "operations": [
    {"operation": "deleteRows", "sheet": "Sheet1", "range": "A50:A100"}
  ]
}
\`\`\``;

// Valid: setHyperlink
export const validSetHyperlink = `\`\`\`json
{
  "plan": "Add link",
  "operations": [
    {"operation": "setHyperlink", "sheet": "Sheet1", "range": "A1", "link": "https://example.com"}
  ]
}
\`\`\``;

// Valid: setCellProtection
export const validSetProtection = `\`\`\`json
{
  "plan": "Protect cells",
  "operations": [
    {"operation": "setCellProtection", "sheet": "Sheet1", "range": "A1", "locked": true, "hidden": false}
  ]
}
\`\`\``;

// Valid: protectWorksheet
export const validProtectSheet = `\`\`\`json
{
  "plan": "Protect sheet",
  "operations": [
    {"operation": "protectWorksheet", "name": "Sheet1", "password": "secret"}
  ]
}
\`\`\``;

// Valid: setPageLayout
export const validPageLayout = `\`\`\`json
{
  "plan": "Set page layout",
  "operations": [
    {"operation": "setPageLayout", "sheet": "Sheet1", "orientation": "landscape", "pageSize": "A4", "margins": {"top": 1, "bottom": 1, "left": 1, "right": 1}}
  ]
}
\`\`\``;

// Valid: addPageBreak
export const validPageBreak = `\`\`\`json
{
  "plan": "Add page break",
  "operations": [
    {"operation": "addPageBreak", "sheet": "Sheet1", "position": "A20"}
  ]
}
\`\`\``;

// Valid: setTableShowHeader
export const validTableHeader = `\`\`\`json
{
  "plan": "Show header",
  "operations": [
    {"operation": "setTableShowHeader", "tableName": "SalesTable", "show": true}
  ]
}
\`\`\``;

// Valid: setChartTitle
export const validChartTitle = `\`\`\`json
{
  "plan": "Set chart title",
  "operations": [
    {"operation": "setChartTitle", "sheet": "Sheet1", "name": "Chart1", "title": "Monthly Sales"}
  ]
}
\`\`\``;

// Valid: addColorScale
export const validColorScale = `\`\`\`json
{
  "plan": "Add color scale",
  "operations": [
    {"operation": "addColorScale", "sheet": "Sheet1", "range": "B2:B100", "colors": ["#FF0000", "#FFFF00", "#00FF00"]}
  ]
}
\`\`\``;

// Valid: addIconSet
export const validIconSet = `\`\`\`json
{
  "plan": "Add icon set",
  "operations": [
    {"operation": "addIconSet", "sheet": "Sheet1", "range": "C2:C100", "iconSet": "threeFlags"}
  ]
}
\`\`\``;

// Valid: addTextComparison
export const validTextComparison = `\`\`\`json
{
  "plan": "Add text comparison",
  "operations": [
    {"operation": "addTextComparison", "sheet": "Sheet1", "range": "D2:D100", "textComparison": "containsText", "operator": "containsText", "text": "urgent"}
  ]
}
\`\`\``;

// Valid: addTopBottom
export const validTopBottom = `\`\`\`json
{
  "plan": "Add top/bottom rule",
  "operations": [
    {"operation": "addTopBottom", "sheet": "Sheet1", "range": "B2:B100", "topBottom": 0, "value": 10, "belowAverage": false}
  ]
}
\`\`\``;

// Valid: addPresetCriteria
export const validPresetCriteria = `\`\`\`json
{
  "plan": "Add preset criteria",
  "operations": [
    {"operation": "addPresetCriteria", "sheet": "Sheet1", "range": "B2:B100", "criteria": "aboveAverage"}
  ]
}
\`\`\``;

// Valid: addCustomConditionalFormat
export const validCustomCF = `\`\`\`json
{
  "plan": "Add custom CF",
  "operations": [
    {"operation": "addCustomConditionalFormat", "sheet": "Sheet1", "range": "B2:B100", "formula": "=B2>100"}
  ]
}
\`\`\``;

// Valid: clearConditionalFormats
export const validClearCF = `\`\`\`json
{
  "plan": "Clear CF",
  "operations": [
    {"operation": "clearConditionalFormats", "sheet": "Sheet1", "range": "B2:B100"}
  ]
}
\`\`\``;

// Valid: addTextBox
export const validAddTextBox = `\`\`\`json
{
  "plan": "Add textbox",
  "operations": [
    {"operation": "addTextBox", "sheet": "Sheet1", "name": "Note1", "left": 100, "top": 200, "width": 150, "height": 50, "text": "Important note"}
  ]
}
\`\`\``;

// Valid: setShapeText
export const validShapeText = `\`\`\`json
{
  "plan": "Set shape text",
  "operations": [
    {"operation": "setShapeText", "sheet": "Sheet1", "name": "Box1", "text": "Updated text"}
  ]
}
\`\`\``;

// Valid: deleteShape
export const validDeleteShape = `\`\`\`json
{
  "plan": "Delete shape",
  "operations": [
    {"operation": "deleteShape", "sheet": "Sheet1", "name": "Box1"}
  ]
}
\`\`\``;

// Valid: addNamedRange
export const validNamedRange = `\`\`\`json
{
  "plan": "Add named range",
  "operations": [
    {"operation": "addNamedRange", "name": "SalesData", "definition": "Sheet1!A1:D100"}
  ]
}
\`\`\``;

// Valid: deleteNamedRange
export const validDeleteNamedRange = `\`\`\`json
{
  "plan": "Delete named range",
  "operations": [
    {"operation": "deleteNamedRange", "name": "SalesData"}
  ]
}
\`\`\``;

// Valid: setFreezePanes
export const validFreezePanes = `\`\`\`json
{
  "plan": "Freeze panes",
  "operations": [
    {"operation": "setFreezePanes", "sheet": "Sheet1", "freezeType": "rows", "row": 1, "column": 0}
  ]
}
\`\`\``;

// Valid: setPrintSettings
export const validPrintSettings = `\`\`\`json
{
  "plan": "Set print",
  "operations": [
    {"operation": "setPrintSettings", "sheet": "Sheet1", "printArea": "A1:D100", "printTitles": "A1:B1"}
  ]
}
\`\`\``;

// Valid: unprotectWorksheet
export const validUnprotect = `\`\`\`json
{
  "plan": "Unprotect",
  "operations": [
    {"operation": "unprotectWorksheet", "name": "Sheet1", "password": "secret"}
  ]
}
\`\`\``;

// Valid: addPivotHierarchy
export const validPivotHierarchy = `\`\`\`json
{
  "plan": "Add pivot hierarchy",
  "operations": [
    {"operation": "addPivotHierarchy", "pivotTable": "Pivot1", "field": "Region", "category": "rows", "index": 0}
  ]
}
\`\`\``;

// Valid: setPivotAggregation
export const validPivotAgg = `\`\`\`json
{
  "plan": "Set pivot aggregation",
  "operations": [
    {"operation": "setPivotAggregation", "pivotTable": "Pivot1", "field": "Sales", "aggregation": "sum"}
  ]
}
\`\`\``;

// Valid: setPivotLayout
export const validPivotLayout = `\`\`\`json
{
  "plan": "Set pivot layout",
  "operations": [
    {"operation": "setPivotLayout", "pivotTable": "Pivot1", "layout": "tabular"}
  ]
}
\`\`\``;

// Valid: addPivotFilter
export const validPivotFilter = `\`\`\`json
{
  "plan": "Add pivot filter",
  "operations": [
    {"operation": "addPivotFilter", "pivotTable": "Pivot1", "field": "Year", "criterion": "2024"}
  ]
}
\`\`\``;

// Valid: deletePivotTable
export const validDeletePivot = `\`\`\`json
{
  "plan": "Delete pivot",
  "operations": [
    {"operation": "deletePivotTable", "name": "Pivot1"}
  ]
}
\`\`\``;

// Valid: addSlicer
export const validAddSlicer = `\`\`\`json
{
  "plan": "Add slicer",
  "operations": [
    {"operation": "addSlicer", "pivotTable": "Pivot1", "field": "Region", "slicerName": "RegionSlicer", "sheet": "PivotSheet"}
  ]
}
\`\`\``;

// Valid: deleteSlicer
export const validDeleteSlicer = `\`\`\`json
{
  "plan": "Delete slicer",
  "operations": [
    {"operation": "deleteSlicer", "sheet": "PivotSheet", "name": "RegionSlicer"}
  ]
}
\`\`\``;

// Valid: deleteChart
export const validDeleteChart = `\`\`\`json
{
  "plan": "Delete chart",
  "operations": [
    {"operation": "deleteChart", "sheet": "Sheet1", "name": "Chart1"}
  ]
}
\`\`\``;

// Valid: setChartSeries
export const validChartSeries = `\`\`\`json
{
  "plan": "Set chart series",
  "operations": [
    {"operation": "setChartSeries", "sheet": "Sheet1", "name": "Chart1", "dataRange": "A1:B10", "seriesName": "Sales"}
  ]
}
\`\`\``;

// Valid: setChartPosition
export const validChartPosition = `\`\`\`json
{
  "plan": "Set chart position",
  "operations": [
    {"operation": "setChartPosition", "sheet": "Sheet1", "name": "Chart1", "left": 200, "top": 100, "width": 400, "height": 300}
  ]
}
\`\`\``;

// Valid: setChartShowLegend
export const validChartLegend = `\`\`\`json
{
  "plan": "Show legend",
  "operations": [
    {"operation": "setChartShowLegend", "sheet": "Sheet1", "name": "Chart1", "show": true}
  ]
}
\`\`\``;

// Valid: setChartShowDataLabels
export const validChartDataLabels = `\`\`\`json
{
  "plan": "Show data labels",
  "operations": [
    {"operation": "setChartShowDataLabels", "sheet": "Sheet1", "name": "Chart1", "show": true}
  ]
}
\`\`\``;

// Valid: deleteTable
export const validDeleteTable = `\`\`\`json
{
  "plan": "Delete table",
  "operations": [
    {"operation": "deleteTable", "name": "SalesTable"}
  ]
}
\`\`\``;

// Valid: addTableColumn
export const validAddTableCol = `\`\`\`json
{
  "plan": "Add column",
  "operations": [
    {"operation": "addTableColumn", "tableName": "SalesTable", "name": "Status"}
  ]
}
\`\`\``;

// Valid: deleteTableRow
export const validDeleteTableRow = `\`\`\`json
{
  "plan": "Delete row",
  "operations": [
    {"operation": "deleteTableRow", "tableName": "SalesTable", "index": 5}
  ]
}
\`\`\``;

// Valid: deleteTableColumn
export const validDeleteTableCol = `\`\`\`json
{
  "plan": "Delete column",
  "operations": [
    {"operation": "deleteTableColumn", "tableName": "SalesTable", "index": 3}
  ]
}
\`\`\``;

// Valid: sortTable
export const validSortTable = `\`\`\`json
{
  "plan": "Sort table",
  "operations": [
    {"operation": "sortTable", "tableName": "SalesTable", "columns": [{"columnIndex": 1, "order": "ascending"}]}
  ]
}
\`\`\``;

// Valid: filterTable
export const validFilterTable = `\`\`\`json
{
  "plan": "Filter table",
  "operations": [
    {"operation": "filterTable", "tableName": "SalesTable", "columnIndex": 2, "rule": "containsText"}
  ]
}
\`\`\``;

// Valid: setTableStyle
export const validTableStyle = `\`\`\`json
{
  "plan": "Set table style",
  "operations": [
    {"operation": "setTableStyle", "tableName": "SalesTable", "style": "TableStyleMedium1"}
  ]
}
\`\`\``;

// Valid: setTableShowTotals
export const validTableTotals = `\`\`\`json
{
  "plan": "Show totals",
  "operations": [
    {"operation": "setTableShowTotals", "tableName": "SalesTable", "show": true}
  ]
}
\`\`\``;

// Valid: setTableBandedRows
export const validTableBandedRows = `\`\`\`json
{
  "plan": "Banded rows",
  "operations": [
    {"operation": "setTableBandedRows", "tableName": "SalesTable", "banded": true}
  ]
}
\`\`\``;

// Valid: setTableBandedColumns
export const validTableBandedCols = `\`\`\`json
{
  "plan": "Banded columns",
  "operations": [
    {"operation": "setTableBandedColumns", "tableName": "SalesTable", "banded": true}
  ]
}
\`\`\``;

// Valid: deleteComment
export const validDeleteComment = `\`\`\`json
{
  "plan": "Delete comment",
  "operations": [
    {"operation": "deleteComment", "sheet": "Sheet1", "range": "A1"}
  ]
}
\`\`\``;

// Valid: addNote
export const validAddNote = `\`\`\`json
{
  "plan": "Add note",
  "operations": [
    {"operation": "addNote", "sheet": "Sheet1", "range": "A1", "text": "Note text"}
  ]
}
\`\`\``;

// Valid: activateWorksheet
export const validActivateSheet = `\`\`\`json
{
  "plan": "Activate sheet",
  "operations": [
    {"operation": "activateWorksheet", "name": "Sheet1"}
  ]
}
\`\`\``;

// Valid: showWorksheet
export const validShowSheet = `\`\`\`json
{
  "plan": "Show sheet",
  "operations": [
    {"operation": "showWorksheet", "name": "Sheet2"}
  ]
}
\`\`\``;

// Valid: setFontItalic
export const validSetFontItalic = `\`\`\`json
{
  "plan": "Italic",
  "operations": [
    {"operation": "setFontItalic", "sheet": "Sheet1", "range": "A1", "italic": true}
  ]
}
\`\`\``;

// Valid: setFontUnderline
export const validSetFontUnderline = `\`\`\`json
{
  "plan": "Underline",
  "operations": [
    {"operation": "setFontUnderline", "sheet": "Sheet1", "range": "A1", "underline": true}
  ]
}
\`\`\``;

// Valid: setFontStrikethrough
export const validSetFontStrike = `\`\`\`json
{
  "plan": "Strikethrough",
  "operations": [
    {"operation": "setFontStrikethrough", "sheet": "Sheet1", "range": "A1", "strikethrough": true}
  ]
}
\`\`\``;

// Valid: setFontColor
export const validSetFontColor = `\`\`\`json
{
  "plan": "Font color",
  "operations": [
    {"operation": "setFontColor", "sheet": "Sheet1", "range": "A1", "color": "#FF0000"}
  ]
}
\`\`\``;

// Valid: setVerticalAlignment
export const validSetVAlign = `\`\`\`json
{
  "plan": "Vertical align",
  "operations": [
    {"operation": "setVerticalAlignment", "sheet": "Sheet1", "range": "A1", "alignment": "center"}
  ]
}
\`\`\``;

// Valid: setTextWrap
export const validSetTextWrap = `\`\`\`json
{
  "plan": "Text wrap",
  "operations": [
    {"operation": "setTextWrap", "sheet": "Sheet1", "range": "A1", "wrap": true}
  ]
}
\`\`\``;

// Valid: setIndentLevel
export const validSetIndent = `\`\`\`json
{
  "plan": "Indent",
  "operations": [
    {"operation": "setIndentLevel", "sheet": "Sheet1", "range": "A1", "level": 2}
  ]
}
\`\`\``;

// Valid: setRotationAngle
export const validSetRotation = `\`\`\`json
{
  "plan": "Rotation",
  "operations": [
    {"operation": "setRotationAngle", "sheet": "Sheet1", "range": "A1", "angle": 45}
  ]
}
\`\`\``;

// Valid: autofitRows
export const validAutofitRows = `\`\`\`json
{
  "plan": "Autofit rows",
  "operations": [
    {"operation": "autofitRows", "sheet": "Sheet1"}
  ]
}
\`\`\``;

// Valid: deleteColumns
export const validDeleteColumns = `\`\`\`json
{
  "plan": "Delete columns",
  "operations": [
    {"operation": "deleteColumns", "sheet": "Sheet1", "range": "E:E"}
  ]
}
\`\`\``;

// Valid: writeFormulas
export const validWriteFormulas = `\`\`\`json
{
  "plan": "Write formulas",
  "operations": [
    {"operation": "writeFormulas", "sheet": "Sheet1", "range": "C2:C10", "formulas": [["=SUM(A2:B2)"]]}
  ]
}
\`\`\``;

export default {
  validStandardBlock,
  validGenericBlock,
  validPlainText,
  validExtraText,
  validNoNewline,
  invalidMalformedJSON,
  invalidMissingFields,
  invalidOpsNotArray,
  invalidEmpty,
  invalidNull,
  validEmptyOps,
  invalidUnknownOp,
  invalidMissingField,
  invalidWrongType,
  validMultiOps,
  validCompletionClaim,
  validCompletionText,
  validSwitchMode,
  invalidJSONWithComments,
  invalidNullValues,
  validLongResponse,
  invalidMultipleBlocks,
  validSetBorder,
  validAddChart,
  validCreatePivot,
  validSetNumberFormat,
  validSetHAlign,
  validSetMerge,
  validSetFontSize,
  validSetFontName,
  validAddDataBar,
  validAddComment,
  validAddShape,
  validAddValidation,
  validRenameSheet,
  validHideSheet,
  validCopySheet,
  validSortRange,
  validCreateTable,
  validAddTableRow,
  validDeleteRows,
  validSetHyperlink,
  validSetProtection,
  validProtectSheet,
  validPageLayout,
  validPageBreak,
  validTableHeader,
  validChartTitle,
  validColorScale,
  validIconSet,
  validTextComparison,
  validTopBottom,
  validPresetCriteria,
  validCustomCF,
  validClearCF,
  validAddTextBox,
  validShapeText,
  validDeleteShape,
  validNamedRange,
  validDeleteNamedRange,
  validFreezePanes,
  validPrintSettings,
  validUnprotect,
  validPivotHierarchy,
  validPivotAgg,
  validPivotLayout,
  validPivotFilter,
  validDeletePivot,
  validAddSlicer,
  validDeleteSlicer,
  validDeleteChart,
  validChartSeries,
  validChartPosition,
  validChartLegend,
  validChartDataLabels,
  validDeleteTable,
  validAddTableCol,
  validDeleteTableRow,
  validDeleteTableCol,
  validSortTable,
  validFilterTable,
  validTableStyle,
  validTableTotals,
  validTableBandedRows,
  validTableBandedCols,
  validDeleteComment,
  validAddNote,
  validActivateSheet,
  validShowSheet,
  validSetFontItalic,
  validSetFontUnderline,
  validSetFontStrike,
  validSetFontColor,
  validSetVAlign,
  validSetTextWrap,
  validSetIndent,
  validSetRotation,
  validAutofitRows,
  validDeleteColumns,
  validWriteFormulas,
};
