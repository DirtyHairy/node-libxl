var xl = require('../lib/libxl'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('FilterColumn', () => {
    let book, sheet, autoFilter, filterColumn;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        sheet = book.addSheet('Sheet1');

        autoFilter = sheet.autoFilter();
        autoFilter.setRef(1, 2, 3, 4);

        filterColumn = autoFilter.column(0);
    });

    it('index returns the column index', () => {
        shouldThrow(filterColumn.index, {});

        expect(filterColumn.index()).toBe(0);
    });

    it('filterType gets the filter type', () => {
        shouldThrow(filterColumn.filterType, {});

        expect(filterColumn.filterType()).toBe(xl.FILTER_NOT_SET);
    });

    it('addFilter and filter set and get a value filter', () => {
        shouldThrow(filterColumn.addFilter, filterColumn, 1);
        shouldThrow(filterColumn.addFilter, {}, 'a');

        expect(filterColumn.addFilter('a')).toBe(filterColumn);

        shouldThrow(filterColumn.filter, filterColumn, 'a');
        shouldThrow(filterColumn.filter, {}, 0);

        expect(filterColumn.filter(0)).toBe('a');
    });

    it('filterSize returns the number of filter values', () => {
        shouldThrow(filterColumn.filterSize, {});

        expect(filterColumn.filterSize()).toBe(0);

        filterColumn.addFilter('a');

        expect(filterColumn.filterSize()).toBe(1);
    });

    it('setTop10 and getTop10 get and set the number of top or bottom items', () => {
        shouldThrow(filterColumn.setTop10, filterColumn, 'a');
        shouldThrow(filterColumn.setTop10, filterColumn, 1, 'a');
        shouldThrow(filterColumn.setTop10, filterColumn, 1, false, 'a');
        shouldThrow(filterColumn.setTop10, {}, 1);
        shouldThrow(filterColumn.setTop10, {}, 1, false);
        shouldThrow(filterColumn.setTop10, {}, 1, false, false);

        expect(filterColumn.setTop10(1)).toBe(filterColumn);
        expect(filterColumn.setTop10(1, false)).toBe(filterColumn);
        expect(filterColumn.setTop10(1, false, false)).toBe(filterColumn);

        shouldThrow(filterColumn.getTop10, {});

        expect(filterColumn.getTop10()).toEqual({ value: 1, top: false, percent: false });
    });

    it('setCustomFilter and getCustomFilter get and set a custom filter', () => {
        shouldThrow(filterColumn.setCustomFilter, filterColumn, xl.OPERATOR_EQUAL, 1);
        shouldThrow(filterColumn.setCustomFilter, filterColumn, xl.OPERATOR_EQUAL, 'a', xl.OPERATOR_EQUAL, 'b', 1);

        expect(filterColumn.setCustomFilter(xl.OPERATOR_EQUAL, 'a')).toBe(filterColumn);
        expect(filterColumn.setCustomFilter(xl.OPERATOR_EQUAL, 'a', xl.OPERATOR_EQUAL, 'b', true)).toBe(filterColumn);

        shouldThrow(filterColumn.getCustomFilter, {});

        expect(filterColumn.getCustomFilter()).toEqual({
            op1: xl.OPERATOR_EQUAL,
            v1: 'a',
            op2: xl.OPERATOR_EQUAL,
            v2: 'b',
            andOp: true,
        });
    });
});
