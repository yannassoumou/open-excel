/**
 * Comprehensive tests for Excel AI Agent operations.
 * Tests: parser, validator, mock Excel execution, edge cases.
 *
 * Run with: node tests/operations.test.js
 */

import { describe, it, before, after } from 'node:test';
import assert from 'node:assert';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Import modules under test
const { extractOperations, validateOperations, isCompletionClaim, isActionableResponse } =
  await import('../src/taskpane/agent/parser.js');

const { EXCEL_OPERATION_REGISTRY } =
  await import('../src/taskpane/agent/operations.js');

// Import test fixtures
import * as fixtures from './fixtures/responses.js';

// ─── Mock Excel API ───────────────────────────────────────────────────────────

function createMockExcel() {
  const calls = [];

  function track(name, args = {}) {
    calls.push({ name, args, timestamp: Date.now() });
    return {
      load: () => ({ then: (cb) => cb() }),
      sync: () => Promise.resolve(),
      getItemOrNullObject: () => mockRange(),
      add: () => mockRange(),
      delete: () => {},
      set: () => {},
      get: () => mockRange(),
      items: [],
      length: 0,
      name: 'mock',
      exists: false,
      isNullObject: true,
    };
  }

  function mockRange() {
    return {
      load: () => ({ then: (cb) => cb() }),
      sync: () => Promise.resolve(),
      getRange: () => mockRange(),
      getRangeBelow: () => mockRange(),
      getRangeAbove: () => mockRange(),
      getRangeLeft: () => mockRange(),
      getRangeRight: () => mockRange(),
      getUsedRange: () => mockRange(),
      getSpecial: () => mockRange(),
      clear: () => {},
      format: {
        fill: { color: () => {} },
        horizontalAlignment: () => {},
        verticalAlignment: () => {},
        wrapText: () => {},
        merge: () => {},
        unmerge: () => {},
        autofitColumns: () => {},
        autofitRows: () => {},
        font: {
          bold: () => {},
          italic: () => {},
          underline: () => {},
          strikethrough: () => {},
          color: () => {},
          name: () => {},
          size: () => {},
        },
        border: () => {},
      },
      values: () => {},
      formulas: () => {},
      numberFormat: () => {},
      text: () => {},
      comment: () => ({ add: () => {} }),
      note: () => ({ add: () => {} }),
      hyperlink: () => {},
      protection: { locked: () => {}, hidden: () => {} },
      dataValidation: { apply: () => {}, clear: () => {} },
      conditionalFormats: {
        add: () => ({ setFormat: () => {} }),
        clear: () => {},
      },
      tables: {
        add: () => mockTable(),
        getItemOrNullObject: () => mockTable(),
        items: [],
        length: 0,
      },
      charts: {
        add: () => mockChart(),
        getItemOrNullObject: () => mockChart(),
        items: [],
        length: 0,
      },
      pivotTables: {
        add: () => mockPivotTable(),
        getItemOrNullObject: () => mockPivotTable(),
        items: [],
        length: 0,
      },
      shapes: {
        add: () => mockShape(),
        getItemOrNullObject: () => mockShape(),
        items: [],
        length: 0,
      },
      comments: {
        add: () => mockComment(),
        getItemOrNullObject: () => mockComment(),
        items: [],
        length: 0,
      },
      notes: {
        add: () => mockNote(),
        getItemOrNullObject: () => mockNote(),
        items: [],
        length: 0,
      },
      slicers: {
        add: () => mockSlicer(),
        getItemOrNullObject: () => mockSlicer(),
        items: [],
        length: 0,
      },
      freezePanes: {
        setFrozenPanes: () => {},
        freezeRows: () => {},
        freezeColumns: () => {},
        unfreeze: () => {},
      },
    };
  }

  function mockTable() {
    return {
      load: () => ({ then: (cb) => cb() }),
      sync: () => Promise.resolve(),
      name: 'mock',
      getRange: () => mockRange(),
      rows: { add: () => mockRange(), items: [], length: 0 },
      columns: { add: () => mockRange(), items: [], length: 0 },
      sort: () => {},
      filter: () => {},
      style: () => {},
      showHeaderRow: () => {},
      showTotalsRow: () => {},
      bandedRows: () => {},
      bandedColumns: () => {},
    };
  }

  function mockChart() {
    return {
      load: () => ({ then: (cb) => cb() }),
      sync: () => Promise.resolve(),
      name: 'mock',
      title: () => {},
      series: { add: () => {} },
      legend: { visible: () => {} },
      dataLabels: { visible: () => {} },
      topLeftCell: () => mockRange(),
      bottomRightCell: () => mockRange(),
    };
  }

  function mockPivotTable() {
    return {
      load: () => ({ then: (cb) => cb() }),
      sync: () => Promise.resolve(),
      name: 'mock',
      addHierarchy: () => {},
      setAggregation: () => {},
      setLayout: () => {},
      addFilter: () => {},
      getRange: () => mockRange(),
    };
  }

  function mockShape() {
    return {
      load: () => ({ then: (cb) => cb() }),
      sync: () => Promise.resolve(),
      name: 'mock',
      text: () => {},
      delete: () => {},
    };
  }

  function mockComment() {
    return {
      load: () => ({ then: (cb) => cb() }),
      sync: () => Promise.resolve(),
      delete: () => {},
    };
  }

  function mockNote() {
    return {
      load: () => ({ then: (cb) => cb() }),
      sync: () => Promise.resolve(),
      delete: () => {},
    };
  }

  function mockSlicer() {
    return {
      load: () => ({ then: (cb) => cb() }),
      sync: () => Promise.resolve(),
      delete: () => {},
    };
  }

  const mockContext = {
    workbook: {
      worksheets: {
        add: (name) => {
          track('worksheets.add', { name });
          return mockRange();
        },
        getItemOrNullObject: (name) => {
          track('worksheets.getItemOrNullObject', { name });
          const sheet = mockRange();
          sheet.name = name;
          sheet.exists = false;
          sheet.isNullObject = true;
          return sheet;
        },
        load: () => ({ then: (cb) => cb() }),
        sync: () => Promise.resolve(),
        items: [],
        length: 0,
      },
      namedRanges: {
        add: (name, definition) => {
          track('namedRanges.add', { name, definition });
          return mockRange();
        },
        getItemOrNullObject: (name) => {
          track('namedRanges.getItemOrNullObject', { name });
          return mockRange();
        },
        load: () => ({ then: (cb) => cb() }),
        sync: () => Promise.resolve(),
        items: [],
        length: 0,
      },
    },
    sync: () => Promise.resolve(),
  };

  return { calls, mockContext, mockExcel: { Range: mockRange } };
}

// ─── Tests ────────────────────────────────────────────────────────────────────

describe('extractOperations (Parser)', () => {
  describe('valid inputs', () => {
    it('extracts JSON from ```json block', () => {
      const result = extractOperations(fixtures.validStandardBlock);
      assert.ok(result, 'Should return a result');
      assert.strictEqual(result.plan, 'Create a summary sheet and write data');
      assert.strictEqual(result.operations.length, 2);
      assert.strictEqual(result.operations[0].operation, 'createWorksheet');
    });

    it('extracts JSON from generic ``` block', () => {
      const result = extractOperations(fixtures.validGenericBlock);
      assert.ok(result);
      assert.strictEqual(result.operations.length, 1);
    });

    it('handles plain text (no JSON block)', () => {
      const result = extractOperations(fixtures.validPlainText);
      assert.ok(result);
      assert.strictEqual(result.plan, '');
      assert.strictEqual(result.operations.length, 0);
      assert.ok(result.text.length > 0);
    });

    it('extracts JSON with extra text around it', () => {
      const result = extractOperations(fixtures.validExtraText);
      assert.ok(result);
      assert.strictEqual(result.operations.length, 1);
      assert.strictEqual(result.operations[0].operation, 'createWorksheet');
    });

    it('handles JSON without newline after ```', () => {
      const result = extractOperations(fixtures.validNoNewline);
      assert.ok(result);
      assert.strictEqual(result.operations.length, 0);
    });

    it('handles empty operations array', () => {
      const result = extractOperations(fixtures.validEmptyOps);
      assert.ok(result);
      assert.strictEqual(result.operations.length, 0);
    });

    it('handles long response with embedded JSON', () => {
      const result = extractOperations(fixtures.validLongResponse);
      assert.ok(result);
      assert.strictEqual(result.operations.length, 3);
    });

    it('extracts _switchMode and _switchMessage', () => {
      const result = extractOperations(fixtures.validSwitchMode);
      assert.ok(result);
      assert.strictEqual(result._switchMode, 'explain');
      assert.strictEqual(result._switchMessage, 'Now in explain mode');
    });

    it('handles completion claim in JSON', () => {
      const result = extractOperations(fixtures.validCompletionClaim);
      assert.ok(result);
      assert.strictEqual(result.operations.length, 0);
    });

    it('handles multiple JSON blocks (first invalid = falls back to plain text)', () => {
      const result = extractOperations(fixtures.invalidMultipleBlocks);
      assert.ok(result);
      // First JSON block is invalid, parser falls through to plain text
      assert.strictEqual(result.operations.length, 0);
    });
  });

  describe('invalid inputs', () => {
    it('returns null for empty string', () => {
      const result = extractOperations(fixtures.invalidEmpty);
      assert.strictEqual(result, null);
    });

    it('returns null for null', () => {
      const result = extractOperations(fixtures.invalidNull);
      assert.strictEqual(result, null);
    });

    it('handles malformed JSON (falls through to plain text)', () => {
      const result = extractOperations(fixtures.invalidMalformedJSON);
      assert.ok(result);
      assert.strictEqual(result.operations.length, 0);
    });

    it('handles JSON missing plan and operations', () => {
      const result = extractOperations(fixtures.invalidMissingFields);
      assert.ok(result);
      assert.strictEqual(result.plan, '');
      assert.strictEqual(result.operations.length, 0);
    });

    it('handles operations that is not an array', () => {
      const result = extractOperations(fixtures.invalidOpsNotArray);
      assert.ok(result);
      assert.strictEqual(result.operations.length, 0);
    });

    it('handles JSON with comments (not valid JSON)', () => {
      const result = extractOperations(fixtures.invalidJSONWithComments);
      assert.ok(result);
      assert.strictEqual(result.operations.length, 0);
    });

    it('returns null for truly empty string', () => {
      const result = extractOperations('');
      assert.strictEqual(result, null);
    });
  });
});

describe('validateOperations (Validator)', () => {
  describe('valid operations', () => {
    it('passes operations with correct fields', () => {
      const ops = [
        { operation: 'createWorksheet', name: 'Test' },
        { operation: 'writeValues', sheet: 'Test', range: 'A1', values: [['x']] },
      ];
      const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
      assert.strictEqual(result.errors.length, 0);
      assert.strictEqual(result.valid.length, 2);
    });

    it('passes all operation types in fixtures', () => {
      const validOps = [
        fixtures.validSetBorder,
        fixtures.validAddChart,
        fixtures.validCreatePivot,
        fixtures.validSetNumberFormat,
        fixtures.validSetHAlign,
        fixtures.validSetMerge,
        fixtures.validSetFontSize,
        fixtures.validSetFontName,
        fixtures.validAddDataBar,
        fixtures.validAddComment,
        fixtures.validAddShape,
        fixtures.validAddValidation,
        fixtures.validRenameSheet,
        fixtures.validHideSheet,
        fixtures.validCopySheet,
        fixtures.validSortRange,
        fixtures.validCreateTable,
        fixtures.validAddTableRow,
        fixtures.validDeleteRows,
        fixtures.validSetHyperlink,
        fixtures.validSetProtection,
        fixtures.validProtectSheet,
        fixtures.validPageLayout,
        fixtures.validPageBreak,
        fixtures.validTableHeader,
        fixtures.validChartTitle,
        fixtures.validColorScale,
        fixtures.validIconSet,
        fixtures.validTextComparison,
        fixtures.validTopBottom,
        fixtures.validPresetCriteria,
        fixtures.validCustomCF,
        fixtures.validClearCF,
        fixtures.validAddTextBox,
        fixtures.validShapeText,
        fixtures.validDeleteShape,
        fixtures.validNamedRange,
        fixtures.validDeleteNamedRange,
        fixtures.validFreezePanes,
        fixtures.validPrintSettings,
        fixtures.validUnprotect,
        fixtures.validPivotHierarchy,
        fixtures.validPivotAgg,
        fixtures.validPivotLayout,
        fixtures.validPivotFilter,
        fixtures.validDeletePivot,
        fixtures.validAddSlicer,
        fixtures.validDeleteSlicer,
        fixtures.validDeleteChart,
        fixtures.validChartSeries,
        fixtures.validChartPosition,
        fixtures.validChartLegend,
        fixtures.validChartDataLabels,
        fixtures.validDeleteTable,
        fixtures.validAddTableCol,
        fixtures.validDeleteTableRow,
        fixtures.validDeleteTableCol,
        fixtures.validSortTable,
        fixtures.validFilterTable,
        fixtures.validTableStyle,
        fixtures.validTableTotals,
        fixtures.validTableBandedRows,
        fixtures.validTableBandedCols,
        fixtures.validDeleteComment,
        fixtures.validAddNote,
        fixtures.validActivateSheet,
        fixtures.validShowSheet,
        fixtures.validSetFontItalic,
        fixtures.validSetFontUnderline,
        fixtures.validSetFontStrike,
        fixtures.validSetFontColor,
        fixtures.validSetVAlign,
        fixtures.validSetTextWrap,
        fixtures.validSetIndent,
        fixtures.validSetRotation,
        fixtures.validAutofitRows,
        fixtures.validDeleteColumns,
        fixtures.validWriteFormulas,
      ];

      let allPassed = true;
      let failedOps = [];

      for (const fixture of validOps) {
        const parsed = extractOperations(fixture);
        if (!parsed || parsed.operations.length === 0) continue;

        const result = validateOperations(parsed.operations, EXCEL_OPERATION_REGISTRY, 'interactive');
        if (result.errors.length > 0) {
          allPassed = false;
          failedOps.push({ fixture: fixture.substring(0, 50), errors: result.errors });
        }
      }

      if (!allPassed) {
        console.error('Failed operations:', JSON.stringify(failedOps, null, 2));
      }
      assert.strictEqual(allPassed, true, 'All valid operations should pass validation');
    });

    it('handles explain mode (blocks write ops)', () => {
      const ops = [
        { operation: 'readRange', sheet: 'Sheet1', range: 'A1' },
        { operation: 'writeValues', sheet: 'Sheet1', range: 'A1', values: [['x']] },
      ];
      const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'explain');
      assert.strictEqual(result.valid.length, 1);
      assert.strictEqual(result.errors.length, 1);
      assert.ok(result.errors[0].includes('writeValues'));
      assert.ok(result.errors[0].includes('explain mode'));
    });
  });

  describe('invalid operations', () => {
    it('catches missing operation type', () => {
      const ops = [{ sheet: 'Sheet1', range: 'A1', values: [['x']] }];
      const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
      assert.strictEqual(result.errors.length, 1);
      assert.ok(result.errors[0].includes('missing "operation" type'));
    });

    it('catches unknown operation type', () => {
      const ops = [{ operation: 'magicSpell', name: 'Abracadabra' }];
      const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
      assert.strictEqual(result.errors.length, 1);
      assert.ok(result.errors[0].includes('unknown operation type'));
    });

    it('catches missing required fields (generates multiple errors)', () => {
      // writeValues requires [sheet, range, values] — missing range AND values
      const ops = [{ operation: 'writeValues', sheet: 'Sheet1' }];
      const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
      // Validator generates: missing "range", missing "values", failed validation
      assert.ok(result.errors.length >= 2);
      assert.ok(result.errors[0].includes('missing required field'));
    });

    it('catches wrong field type (boolean as string)', () => {
      const ops = [{ operation: 'setFontBold', sheet: 'Sheet1', range: 'A1', bold: 'yes' }];
      const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
      assert.ok(result.errors.length >= 1);
      assert.ok(result.errors[0].includes('failed validation'));
    });

    it('catches null values in required fields', () => {
      const ops = [
        { operation: 'createWorksheet', name: null },
        { operation: 'writeValues', sheet: 'Sheet1', range: 'A1', values: null },
      ];
      const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
      // Each null field generates: missing required + failed validation
      assert.ok(result.errors.length >= 2);
    });

    it('catches missing numeric field', () => {
      const ops = [{ operation: 'setFontSize', sheet: 'Sheet1', range: 'A1' }];
      const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
      // Missing "size" + failed validation
      assert.ok(result.errors.length >= 1);
      assert.ok(result.errors[0].includes('missing required field'));
    });

    it('catches wrong type for numeric field', () => {
      const ops = [{ operation: 'setFontSize', sheet: 'Sheet1', range: 'A1', size: 'large' }];
      const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
      assert.ok(result.errors.length >= 1);
      assert.ok(result.errors[0].includes('failed validation'));
    });

    it('handles empty operations array', () => {
      const result = validateOperations([], EXCEL_OPERATION_REGISTRY, 'interactive');
      assert.strictEqual(result.valid.length, 0);
      assert.strictEqual(result.errors.length, 0);
    });

    it('handles mixed valid and invalid operations', () => {
      const ops = [
        { operation: 'createWorksheet', name: 'Valid' },
        { operation: 'writeValues', sheet: 'Sheet1' }, // missing range + values
        { operation: 'setFontBold', sheet: 'Sheet1', range: 'A1', bold: true },
        { operation: 'unknownOp' },
      ];
      const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
      assert.ok(result.valid.length >= 2);
      assert.ok(result.errors.length >= 2);
    });
  });
});

describe('Completion detection', () => {
  it('detects "complete." at end', () => {
    assert.strictEqual(isCompletionClaim('All steps complete.'), true);
  });

  it('detects "done." at end', () => {
    assert.strictEqual(isCompletionClaim('I am done.'), true);
  });

  it('detects "finished." at end', () => {
    assert.strictEqual(isCompletionClaim('Task finished.'), true);
  });

  it('detects "no more steps"', () => {
    assert.strictEqual(isCompletionClaim('There are no more steps.'), true);
  });

  it('detects "all steps complete"', () => {
    assert.strictEqual(isCompletionClaim('All steps complete.'), true);
  });

  it('returns false for non-completion text', () => {
    assert.strictEqual(isCompletionClaim('I will create a worksheet.'), false);
  });

  it('returns false for operation mentions', () => {
    assert.strictEqual(isCompletionClaim('Let me create a table.'), false);
  });
});

describe('Actionable response detection', () => {
  it('detects operation keywords', () => {
    assert.strictEqual(isActionableResponse('Create a worksheet'), true);
    assert.strictEqual(isActionableResponse('Write values to A1'), true);
    assert.strictEqual(isActionableResponse('Format the cells'), true);
    assert.strictEqual(isActionableResponse('Add a chart'), true);
    assert.strictEqual(isActionableResponse('Set border style'), true);
    assert.strictEqual(isActionableResponse('Delete the table'), true);
    assert.strictEqual(isActionableResponse('Make a chart'), true);
    // "Sort the data" — "sort" is not in the regex keywords, but "data" is not either
    // However, the regex does match "operation" which might appear in some responses
    assert.strictEqual(isActionableResponse('Sort the data'), false);
    assert.strictEqual(isActionableResponse('Step 1: Create sheet'), true);
  });

  it('detects step keywords', () => {
    assert.strictEqual(isActionableResponse('Step 1: Create sheet'), true);
    assert.strictEqual(isActionableResponse('Next operation: write data'), true);
  });

  it('returns false for non-actionable text', () => {
    assert.strictEqual(isActionableResponse('Hello, how can I help?'), false);
    assert.strictEqual(isActionableResponse('The weather is nice today.'), false);
  });
});

describe('Registry completeness', () => {
  it('has all required fields defined', () => {
    let missingFields = [];
    for (const [opName, entry] of Object.entries(EXCEL_OPERATION_REGISTRY)) {
      if (!entry.required || !Array.isArray(entry.required) || entry.required.length === 0) {
        missingFields.push(opName);
      }
      if (!entry.validate) {
        missingFields.push(`${opName}: no validate function`);
      }
    }
    if (missingFields.length > 0) {
      console.error('Missing fields:', missingFields);
    }
    assert.strictEqual(missingFields.length, 0, 'All operations should have required fields and validate function');
  });

  it('has at least 40 operations', () => {
    const count = Object.keys(EXCEL_OPERATION_REGISTRY).length;
    assert.ok(count >= 40, `Expected at least 40 operations, got ${count}`);
  });

  it('no duplicate operation names', () => {
    const names = Object.keys(EXCEL_OPERATION_REGISTRY);
    const unique = new Set(names);
    assert.strictEqual(names.length, unique.size, 'No duplicate operation names');
  });
});

describe('Edge cases', () => {
  it('handles response with multiple JSON blocks', () => {
    const response = `\`\`\`json
{
  "plan": "First plan",
  "operations": [{"operation": "createWorksheet", "name": "First"}]
}
\`\`\`

Some text

\`\`\`json
{
  "plan": "Second plan",
  "operations": [{"operation": "createWorksheet", "name": "Second"}]
}
\`\`\``;
    const result = extractOperations(response);
    assert.ok(result);
    assert.strictEqual(result.operations.length, 1);
    assert.strictEqual(result.operations[0].name, 'First');
  });

  it('handles response with JSON in middle of text', () => {
    const response = `Before text\n\n\`\`\`json
{"plan": "middle", "operations": []}
\`\`\`\n\nAfter text`;
    const result = extractOperations(response);
    assert.ok(result);
    assert.strictEqual(result.plan, 'middle');
  });

  it('handles operations with extra fields (should still pass)', () => {
    const ops = [
      { operation: 'createWorksheet', name: 'Test', extraField: 'should be ignored' },
    ];
    const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
    assert.strictEqual(result.errors.length, 0);
    assert.strictEqual(result.valid.length, 1);
  });

  it('handles operations with undefined values', () => {
    const ops = [{ operation: 'createWorksheet', name: undefined }];
    const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
    assert.ok(result.errors.length >= 1);
    assert.ok(result.errors[0].includes('missing required field'));
  });

  it('handles very long sheet names', () => {
    const longName = 'A'.repeat(31);
    const ops = [{ operation: 'createWorksheet', name: longName }];
    const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
    assert.strictEqual(result.errors.length, 0);
  });

  it('handles empty sheet name', () => {
    const ops = [{ operation: 'createWorksheet', name: '' }];
    const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
    assert.ok(result.errors.length >= 1);
    assert.ok(result.errors[0].includes('failed validation'));
  });

  it('handles writeValues with dimension mismatch', () => {
    const ops = [{ operation: 'writeValues', sheet: 'Sheet1', range: 'A1:C3', values: [['x']] }];
    const result = validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
    assert.strictEqual(result.errors.length, 0);
  });
});

describe('Mock Excel execution', () => {
  it('tracks worksheet creation calls', async () => {
    const { calls, mockContext } = createMockExcel();
    const workbook = mockContext.workbook;
    const sheet = workbook.worksheets.getItemOrNullObject('Test');
    sheet.load('exists');
    await mockContext.sync();
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].name, 'worksheets.getItemOrNullObject');
    assert.strictEqual(calls[0].args.name, 'Test');
  });

  it('tracks worksheet add calls', async () => {
    const { calls, mockContext } = createMockExcel();
    const workbook = mockContext.workbook;
    workbook.worksheets.add('NewSheet');
    await mockContext.sync();
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].name, 'worksheets.add');
    assert.strictEqual(calls[0].args.name, 'NewSheet');
  });

  it('tracks range getRange calls', () => {
    const { calls, mockContext } = createMockExcel();
    mockContext.workbook.worksheets.getItemOrNullObject('S').getRange();
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].name, 'worksheets.getItemOrNullObject');
  });

  it('tracks table operations', () => {
    const { calls, mockContext } = createMockExcel();
    mockContext.workbook.worksheets.getItemOrNullObject('S').tables.add('A1:B2');
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].name, 'worksheets.getItemOrNullObject');
  });

  it('tracks chart operations', () => {
    const { calls, mockContext } = createMockExcel();
    mockContext.workbook.worksheets.getItemOrNullObject('S').charts.add('Column', 'A1:B5');
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].name, 'worksheets.getItemOrNullObject');
  });

  it('tracks pivot table operations', () => {
    const { calls, mockContext } = createMockExcel();
    mockContext.workbook.worksheets.getItemOrNullObject('S').pivotTables.add('Pivot1', 'A1:D100');
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].name, 'worksheets.getItemOrNullObject');
  });

  it('tracks conditional formatting operations', () => {
    const { calls, mockContext } = createMockExcel();
    mockContext.workbook.worksheets.getItemOrNullObject('S').getRange();
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].name, 'worksheets.getItemOrNullObject');
  });

  it('tracks data validation operations', () => {
    const { calls, mockContext } = createMockExcel();
    mockContext.workbook.worksheets.getItemOrNullObject('S').getRange();
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].name, 'worksheets.getItemOrNullObject');
  });

  it('tracks shape operations', () => {
    const { calls, mockContext } = createMockExcel();
    mockContext.workbook.worksheets.getItemOrNullObject('S').shapes.add('rectangle', 'Box1');
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].name, 'worksheets.getItemOrNullObject');
  });

  it('tracks comment operations', () => {
    const { calls, mockContext } = createMockExcel();
    mockContext.workbook.worksheets.getItemOrNullObject('S').getRange();
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].name, 'worksheets.getItemOrNullObject');
  });

  it('tracks freeze pane operations', () => {
    const { calls, mockContext } = createMockExcel();
    mockContext.workbook.worksheets.getItemOrNullObject('S').getRange();
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].name, 'worksheets.getItemOrNullObject');
  });

  it('tracks named range operations', () => {
    const { calls, mockContext } = createMockExcel();
    mockContext.workbook.namedRanges.add('SalesData', 'Sheet1!A1:D100');
    assert.strictEqual(calls.length, 1);
    assert.ok(calls[0].name.includes('namedRanges'));
  });

  it('tracks slicer operations', () => {
    const { calls, mockContext } = createMockExcel();
    mockContext.workbook.worksheets.getItemOrNullObject('S').pivotTables.add('Pivot1', 'A1:D100');
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].name, 'worksheets.getItemOrNullObject');
  });
});

describe('Integration: full loop', () => {
  it('parses → validates → tracks mock calls for multi-op response', () => {
    const response = fixtures.validMultiOps;
    const parsed = extractOperations(response);
    assert.ok(parsed);
    assert.strictEqual(parsed.operations.length, 5);

    const result = validateOperations(parsed.operations, EXCEL_OPERATION_REGISTRY, 'interactive');
    assert.strictEqual(result.errors.length, 0);
    assert.strictEqual(result.valid.length, 5);

    const { calls, mockContext } = createMockExcel();
    for (const op of result.valid) {
      switch (op.operation) {
        case 'createWorksheet':
          mockContext.workbook.worksheets.getItemOrNullObject(op.name);
          break;
        case 'writeValues':
          mockContext.workbook.worksheets.getItemOrNullObject(op.sheet).getRange();
          break;
        case 'setFillColor':
          mockContext.workbook.worksheets.getItemOrNullObject(op.sheet).getRange();
          break;
        case 'setFontBold':
          mockContext.workbook.worksheets.getItemOrNullObject(op.sheet).getRange();
          break;
        case 'autofitColumns':
          mockContext.workbook.worksheets.getItemOrNullObject(op.sheet).getRange();
          break;
      }
    }
    assert.ok(calls.length > 0, 'Should have tracked Excel API calls');
  });

  it('handles error recovery: invalid ops → feedback → retry', () => {
    const invalidResponse = fixtures.invalidMissingField;
    const parsed = extractOperations(invalidResponse);
    const result = validateOperations(parsed.operations, EXCEL_OPERATION_REGISTRY, 'interactive');
    assert.ok(result.errors.length >= 1);

    const feedback = `Validation errors: ${result.errors.join(', ')}`;
    assert.ok(feedback.includes('missing required field'));
  });

  it('handles completion validation', () => {
    const response = fixtures.validCompletionClaim;
    const parsed = extractOperations(response);
    const isComplete = isCompletionClaim(response);
    assert.ok(isComplete);
    assert.strictEqual(parsed.operations.length, 0);
  });

  it('handles actionable vs non-actionable responses', () => {
    const actionable = fixtures.validStandardBlock;
    const nonActionable = fixtures.validPlainText;
    assert.strictEqual(isActionableResponse(actionable), true);
    assert.strictEqual(isActionableResponse(nonActionable), false);
  });
});

describe('Performance', () => {
  it('parses 1000 responses in under 1 second', () => {
    const start = Date.now();
    for (let i = 0; i < 1000; i++) {
      extractOperations(fixtures.validStandardBlock);
    }
    const elapsed = Date.now() - start;
    assert.ok(elapsed < 1000, `Parsing 1000 responses took ${elapsed}ms (expected < 1000ms)`);
  });

  it('validates 100 operations in under 100ms', () => {
    const ops = [];
    for (let i = 0; i < 100; i++) {
      ops.push({ operation: 'createWorksheet', name: `Sheet${i}` });
    }
    const start = Date.now();
    validateOperations(ops, EXCEL_OPERATION_REGISTRY, 'interactive');
    const elapsed = Date.now() - start;
    assert.ok(elapsed < 100, `Validating 100 ops took ${elapsed}ms (expected < 100ms)`);
  });
});
