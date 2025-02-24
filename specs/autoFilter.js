var xl = require('../lib/libxl'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('AutoFilter', () => {
    let book, sheet, autoFilter;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        sheet = book.addSheet('Sheet1');
        autoFilter = sheet.autoFilter();
    });

    it('setRef and getRef set and get the range of the autofilter', () => {
        shouldThrow(autoFilter.setRef, autoFilter, 1, 2, 3, 'a');
        shouldThrow(autoFilter.setRef, {}, 1, 2, 3, 4);

        expect(autoFilter.setRef(1, 2, 3, 4)).toBe(autoFilter);

        shouldThrow(autoFilter.getRef, {});

        expect(autoFilter.getRef()).toEqual({ rowFirst: 1, rowLast: 2, colFirst: 3, colLast: 4 });
    });

    it('column gets a filter column', () => {
        autoFilter.setRef(1, 2, 3, 4);

        shouldThrow(autoFilter.column, autoFilter, 'a');
        shouldThrow(autoFilter.column, {}, 0);

        expect(autoFilter.column(0)).toBeInstanceOf(xl.FilterColumn);
    });

    it('columnSize returns the number of filter columns', () => {
        autoFilter.setRef(1, 2, 3, 4);

        shouldThrow(autoFilter.columnSize, {});

        expect(autoFilter.columnSize()).toBe(0);

        autoFilter.column(0);

        expect(autoFilter.columnSize()).toBe(1);
    });

    it('getSortRange return the sort range', () => {
        autoFilter.setRef(1, 2, 3, 4);
        autoFilter.setSort(1);

        shouldThrow(autoFilter.getSortRange, {});

        expect(autoFilter.getSortRange()).toEqual({ rowFirst: 2, rowLast: 2, colFirst: 3, colLast: 4 });
    });

    it('getSort and setSort get and set the sort column', () => {
        autoFilter.setRef(1, 2, 3, 4);

        shouldThrow(autoFilter.setSort, autoFilter, 'a');
        shouldThrow(autoFilter.setSort, autoFilter, 1, 1);
        shouldThrow(autoFilter.setSort, {}, 1);
        shouldThrow(autoFilter.setSort, {}, 1, false);

        expect(autoFilter.setSort(1)).toBe(autoFilter);
        expect(autoFilter.setSort(1, false)).toBe(autoFilter);

        shouldThrow(autoFilter.getSort, autoFilter, 'a');
        shouldThrow(autoFilter.getSort, {});

        expect(autoFilter.getSort()).toEqual({ columnIndex: 1, descending: false });
        expect(autoFilter.getSort(0)).toEqual({ columnIndex: 1, descending: false });
    });

    it('addSort adds a sort column', () => {
        autoFilter.setRef(1, 2, 3, 4);

        shouldThrow(autoFilter.addSort, autoFilter, 'a');
        shouldThrow(autoFilter.addSort, autoFilter, 1, 1);
        shouldThrow(autoFilter.addSort, {}, 1);
        shouldThrow(autoFilter.addSort, {}, 1, false);

        expect(autoFilter.addSort(1)).toBe(autoFilter);
        expect(autoFilter.addSort(1, false)).toBe(autoFilter);
    });
});
