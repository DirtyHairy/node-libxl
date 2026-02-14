const { describe, it, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
var xl = require('../lib/libxl'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('Table', () => {
    let book, sheet, table;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        sheet = book.addSheet('Sheet1');
        table = sheet.addTable('Table1', 0, 5, 0, 3);
    });

    it('name and setName manage the table name', () => {
        shouldThrow(table.name, {});

        assert.strictEqual(table.name(), 'Table1');

        shouldThrow(table.setName, table, 1);
        shouldThrow(table.setName, {}, 'NewName');

        assert.strictEqual(table.setName('NewName'), table);
        assert.strictEqual(table.name(), 'NewName');
    });

    it('ref and setRef manage the table ref', () => {
        shouldThrow(table.ref, {});

        assert.strictEqual(typeof table.ref(), 'string');

        shouldThrow(table.setRef, table, 1);
        shouldThrow(table.setRef, {}, 'A1:D6');

        assert.strictEqual(table.setRef('A1:D10'), table);
    });

    it('autoFilter returns an AutoFilter instance', () => {
        shouldThrow(table.autoFilter, {});

        assert.ok(table.autoFilter() instanceof xl.AutoFilter);
    });

    it('style and setStyle manage the table style', () => {
        shouldThrow(table.style, {});

        assert.strictEqual(typeof table.style(), 'number');

        shouldThrow(table.setStyle, table, 'a');
        shouldThrow(table.setStyle, {}, xl.TABLESTYLE_LIGHT1);

        assert.strictEqual(table.setStyle(xl.TABLESTYLE_LIGHT1), table);
        assert.strictEqual(table.style(), xl.TABLESTYLE_LIGHT1);
    });

    it('showRowStripes and setShowRowStripes manage row stripes', () => {
        shouldThrow(table.showRowStripes, {});

        assert.strictEqual(typeof table.showRowStripes(), 'boolean');

        shouldThrow(table.setShowRowStripes, table, 'a');
        shouldThrow(table.setShowRowStripes, {}, true);

        assert.strictEqual(table.setShowRowStripes(false), table);
        assert.strictEqual(table.showRowStripes(), false);

        table.setShowRowStripes(true);
        assert.strictEqual(table.showRowStripes(), true);
    });

    it('showColumnStripes and setShowColumnStripes manage column stripes', () => {
        shouldThrow(table.showColumnStripes, {});

        assert.strictEqual(typeof table.showColumnStripes(), 'boolean');

        shouldThrow(table.setShowColumnStripes, table, 'a');
        shouldThrow(table.setShowColumnStripes, {}, true);

        assert.strictEqual(table.setShowColumnStripes(true), table);
        assert.strictEqual(table.showColumnStripes(), true);

        table.setShowColumnStripes(false);
        assert.strictEqual(table.showColumnStripes(), false);
    });

    it('showFirstColumn and setShowFirstColumn manage first column highlighting', () => {
        shouldThrow(table.showFirstColumn, {});

        assert.strictEqual(typeof table.showFirstColumn(), 'boolean');

        shouldThrow(table.setShowFirstColumn, table, 'a');
        shouldThrow(table.setShowFirstColumn, {}, true);

        assert.strictEqual(table.setShowFirstColumn(true), table);
        assert.strictEqual(table.showFirstColumn(), true);

        table.setShowFirstColumn(false);
        assert.strictEqual(table.showFirstColumn(), false);
    });

    it('showLastColumn and setShowLastColumn manage last column highlighting', () => {
        shouldThrow(table.showLastColumn, {});

        assert.strictEqual(typeof table.showLastColumn(), 'boolean');

        shouldThrow(table.setShowLastColumn, table, 'a');
        shouldThrow(table.setShowLastColumn, {}, true);

        assert.strictEqual(table.setShowLastColumn(true), table);
        assert.strictEqual(table.showLastColumn(), true);

        table.setShowLastColumn(false);
        assert.strictEqual(table.showLastColumn(), false);
    });

    it('columnSize returns the number of columns', () => {
        shouldThrow(table.columnSize, {});

        assert.strictEqual(table.columnSize(), 4);
    });

    it('columnName and setColumnName manage column names', () => {
        shouldThrow(table.columnName, table, 'a');
        shouldThrow(table.columnName, {}, 0);

        assert.strictEqual(typeof table.columnName(0), 'string');

        shouldThrow(table.setColumnName, table, 'a', 'Name');
        shouldThrow(table.setColumnName, table, 0, 1);
        shouldThrow(table.setColumnName, {}, 0, 'Name');

        assert.strictEqual(table.setColumnName(0, 'MyColumn'), table);
        assert.strictEqual(table.columnName(0), 'MyColumn');
    });
});
