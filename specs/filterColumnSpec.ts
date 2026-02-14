import { describe, it, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import * as xl from '../lib/libxl';

describe('FilterColumn', () => {
    let book: xl.Book, sheet: xl.Sheet, autoFilter: xl.AutoFilter, filterColumn: xl.FilterColumn;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        sheet = book.addSheet('Sheet1');

        autoFilter = sheet.autoFilter();
        autoFilter.setRef(1, 2, 3, 4);

        filterColumn = autoFilter.column(0);
    });

    it('index returns the column index', () => {
        assert.throws(() => (filterColumn.index as any).call({}));

        assert.strictEqual(filterColumn.index(), 0);
    });

    it('filterType gets the filter type', () => {
        assert.throws(() => (filterColumn.filterType as any).call({}));

        assert.strictEqual(filterColumn.filterType(), xl.FILTER_NOT_SET);
    });

    it('addFilter and filter set and get a value filter', () => {
        assert.throws(() => (filterColumn.addFilter as any).call(filterColumn, 1));
        assert.throws(() => (filterColumn.addFilter as any).call({}, 'a'));

        assert.strictEqual(filterColumn.addFilter('a'), filterColumn);

        assert.throws(() => (filterColumn.filter as any).call(filterColumn, 'a'));
        assert.throws(() => (filterColumn.filter as any).call({}, 0));

        assert.strictEqual(filterColumn.filter(0), 'a');
    });

    it('filterSize returns the number of filter values', () => {
        assert.throws(() => (filterColumn.filterSize as any).call({}));

        assert.strictEqual(filterColumn.filterSize(), 0);

        filterColumn.addFilter('a');

        assert.strictEqual(filterColumn.filterSize(), 1);
    });

    it('setTop10 and getTop10 get and set the number of top or bottom items', () => {
        assert.throws(() => (filterColumn.setTop10 as any).call(filterColumn, 'a'));
        assert.throws(() => (filterColumn.setTop10 as any).call(filterColumn, 1, 'a'));
        assert.throws(() => (filterColumn.setTop10 as any).call(filterColumn, 1, false, 'a'));
        assert.throws(() => (filterColumn.setTop10 as any).call({}, 1));
        assert.throws(() => (filterColumn.setTop10 as any).call({}, 1, false));
        assert.throws(() => (filterColumn.setTop10 as any).call({}, 1, false, false));

        assert.strictEqual(filterColumn.setTop10(1), filterColumn);
        assert.strictEqual(filterColumn.setTop10(1, false), filterColumn);
        assert.strictEqual(filterColumn.setTop10(1, false, false), filterColumn);

        assert.throws(() => (filterColumn.getTop10 as any).call({}));

        assert.deepStrictEqual(filterColumn.getTop10(), { value: 1, top: false, percent: false });
    });

    it('setCustomFilter and getCustomFilter get and set a custom filter', () => {
        assert.throws(() => (filterColumn.setCustomFilter as any).call(filterColumn, xl.OPERATOR_EQUAL, 1));
        assert.throws(() =>
            (filterColumn.setCustomFilter as any).call(filterColumn, xl.OPERATOR_EQUAL, 'a', xl.OPERATOR_EQUAL, 'b', 1),
        );

        assert.strictEqual(filterColumn.setCustomFilter(xl.OPERATOR_EQUAL, 'a'), filterColumn);
        assert.strictEqual(
            filterColumn.setCustomFilter(xl.OPERATOR_EQUAL, 'a', xl.OPERATOR_EQUAL, 'b', true),
            filterColumn,
        );

        assert.throws(() => (filterColumn.getCustomFilter as any).call({}));

        assert.deepStrictEqual(filterColumn.getCustomFilter(), {
            op1: xl.OPERATOR_EQUAL,
            v1: 'a',
            op2: xl.OPERATOR_EQUAL,
            v2: 'b',
            andOp: true,
        });
    });
});
