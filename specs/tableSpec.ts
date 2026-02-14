import { describe, it, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import * as xl from '../lib/libxl.js';

describe('Table', () => {
    let book: xl.Book, sheet: xl.Sheet, table: xl.Table;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        sheet = book.addSheet('Sheet1');
        table = sheet.addTable('Table1', 0, 5, 0, 3);
    });

    it('name and setName manage the table name', () => {
        assert.throws(() => (table.name as any).call({}));

        assert.strictEqual(table.name(), 'Table1');

        assert.throws(() => (table.setName as any).call(table, 1));
        assert.throws(() => (table.setName as any).call({}, 'NewName'));

        assert.strictEqual(table.setName('NewName'), table);
        assert.strictEqual(table.name(), 'NewName');
    });

    it('ref and setRef manage the table ref', () => {
        assert.throws(() => (table.ref as any).call({}));

        assert.strictEqual(typeof table.ref(), 'string');

        assert.throws(() => (table.setRef as any).call(table, 1));
        assert.throws(() => (table.setRef as any).call({}, 'A1:D6'));

        assert.strictEqual(table.setRef('A1:D10'), table);
    });

    it('autoFilter returns an AutoFilter instance', () => {
        assert.throws(() => (table.autoFilter as any).call({}));

        assert.ok(table.autoFilter() instanceof xl.AutoFilter);
    });

    it('style and setStyle manage the table style', () => {
        assert.throws(() => (table.style as any).call({}));

        assert.strictEqual(typeof table.style(), 'number');

        assert.throws(() => (table.setStyle as any).call(table, 'a'));
        assert.throws(() => (table.setStyle as any).call({}, xl.TABLESTYLE_LIGHT1));

        assert.strictEqual(table.setStyle(xl.TABLESTYLE_LIGHT1), table);
        assert.strictEqual(table.style(), xl.TABLESTYLE_LIGHT1);
    });

    it('showRowStripes and setShowRowStripes manage row stripes', () => {
        assert.throws(() => (table.showRowStripes as any).call({}));

        assert.strictEqual(typeof table.showRowStripes(), 'boolean');

        assert.throws(() => (table.setShowRowStripes as any).call(table, 'a'));
        assert.throws(() => (table.setShowRowStripes as any).call({}, true));

        assert.strictEqual(table.setShowRowStripes(false), table);
        assert.strictEqual(table.showRowStripes(), false);

        table.setShowRowStripes(true);
        assert.strictEqual(table.showRowStripes(), true);
    });

    it('showColumnStripes and setShowColumnStripes manage column stripes', () => {
        assert.throws(() => (table.showColumnStripes as any).call({}));

        assert.strictEqual(typeof table.showColumnStripes(), 'boolean');

        assert.throws(() => (table.setShowColumnStripes as any).call(table, 'a'));
        assert.throws(() => (table.setShowColumnStripes as any).call({}, true));

        assert.strictEqual(table.setShowColumnStripes(true), table);
        assert.strictEqual(table.showColumnStripes(), true);

        table.setShowColumnStripes(false);
        assert.strictEqual(table.showColumnStripes(), false);
    });

    it('showFirstColumn and setShowFirstColumn manage first column highlighting', () => {
        assert.throws(() => (table.showFirstColumn as any).call({}));

        assert.strictEqual(typeof table.showFirstColumn(), 'boolean');

        assert.throws(() => (table.setShowFirstColumn as any).call(table, 'a'));
        assert.throws(() => (table.setShowFirstColumn as any).call({}, true));

        assert.strictEqual(table.setShowFirstColumn(true), table);
        assert.strictEqual(table.showFirstColumn(), true);

        table.setShowFirstColumn(false);
        assert.strictEqual(table.showFirstColumn(), false);
    });

    it('showLastColumn and setShowLastColumn manage last column highlighting', () => {
        assert.throws(() => (table.showLastColumn as any).call({}));

        assert.strictEqual(typeof table.showLastColumn(), 'boolean');

        assert.throws(() => (table.setShowLastColumn as any).call(table, 'a'));
        assert.throws(() => (table.setShowLastColumn as any).call({}, true));

        assert.strictEqual(table.setShowLastColumn(true), table);
        assert.strictEqual(table.showLastColumn(), true);

        table.setShowLastColumn(false);
        assert.strictEqual(table.showLastColumn(), false);
    });

    it('columnSize returns the number of columns', () => {
        assert.throws(() => (table.columnSize as any).call({}));

        assert.strictEqual(table.columnSize(), 4);
    });

    it('columnName and setColumnName manage column names', () => {
        assert.throws(() => (table.columnName as any).call(table, 'a'));
        assert.throws(() => (table.columnName as any).call({}, 0));

        assert.strictEqual(typeof table.columnName(0), 'string');

        assert.throws(() => (table.setColumnName as any).call(table, 'a', 'Name'));
        assert.throws(() => (table.setColumnName as any).call(table, 0, 1));
        assert.throws(() => (table.setColumnName as any).call({}, 0, 'Name'));

        assert.strictEqual(table.setColumnName(0, 'MyColumn'), table);
        assert.strictEqual(table.columnName(0), 'MyColumn');
    });
});
