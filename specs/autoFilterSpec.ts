import { describe, it, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import * as xl from '../lib/libxl.js';

describe('AutoFilter', () => {
    let book: xl.Book, sheet: xl.Sheet, autoFilter: xl.AutoFilter;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        sheet = book.addSheet('Sheet1');
        autoFilter = sheet.autoFilter();
    });

    it('setRef and getRef set and get the range of the autofilter', () => {
        assert.throws(() => (autoFilter.setRef as any).call(autoFilter, 1, 2, 3, 'a'));
        assert.throws(() => (autoFilter.setRef as any).call({}, 1, 2, 3, 4));

        assert.strictEqual(autoFilter.setRef(1, 2, 3, 4), autoFilter);

        assert.throws(() => (autoFilter.getRef as any).call({}));

        assert.deepStrictEqual(autoFilter.getRef(), { rowFirst: 1, rowLast: 2, colFirst: 3, colLast: 4 });
    });

    it('column gets a filter column', () => {
        autoFilter.setRef(1, 2, 3, 4);

        assert.throws(() => (autoFilter.column as any).call(autoFilter, 'a'));
        assert.throws(() => (autoFilter.column as any).call({}, 0));

        assert.ok(autoFilter.column(0) instanceof xl.FilterColumn);
    });

    it('columnSize returns the number of filter columns', () => {
        autoFilter.setRef(1, 2, 3, 4);

        assert.throws(() => (autoFilter.columnSize as any).call({}));

        assert.strictEqual(autoFilter.columnSize(), 0);

        autoFilter.column(0);

        assert.strictEqual(autoFilter.columnSize(), 1);
    });

    it('getSortRange return the sort range', () => {
        autoFilter.setRef(1, 2, 3, 4);
        autoFilter.setSort(1);

        assert.throws(() => (autoFilter.getSortRange as any).call({}));

        assert.deepStrictEqual(autoFilter.getSortRange(), { rowFirst: 2, rowLast: 2, colFirst: 3, colLast: 4 });
    });

    it('getSort and setSort get and set the sort column', () => {
        autoFilter.setRef(1, 2, 3, 4);

        assert.throws(() => (autoFilter.setSort as any).call(autoFilter, 'a'));
        assert.throws(() => (autoFilter.setSort as any).call(autoFilter, 1, 1));
        assert.throws(() => (autoFilter.setSort as any).call({}, 1));
        assert.throws(() => (autoFilter.setSort as any).call({}, 1, false));

        assert.strictEqual(autoFilter.setSort(1), autoFilter);
        assert.strictEqual(autoFilter.setSort(1, false), autoFilter);

        assert.throws(() => (autoFilter.getSort as any).call(autoFilter, 'a'));
        assert.throws(() => (autoFilter.getSort as any).call({}));

        assert.deepStrictEqual(autoFilter.getSort(), { columnIndex: 1, descending: false });
        assert.deepStrictEqual(autoFilter.getSort(0), { columnIndex: 1, descending: false });
    });

    it('addSort adds a sort column', () => {
        autoFilter.setRef(1, 2, 3, 4);

        assert.throws(() => (autoFilter.addSort as any).call(autoFilter, 'a'));
        assert.throws(() => (autoFilter.addSort as any).call(autoFilter, 1, 1));
        assert.throws(() => (autoFilter.addSort as any).call({}, 1));
        assert.throws(() => (autoFilter.addSort as any).call({}, 1, false));

        assert.strictEqual(autoFilter.addSort(1), autoFilter);
        assert.strictEqual(autoFilter.addSort(1, false), autoFilter);
    });
});
