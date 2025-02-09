var xl = require('../lib/libxl'),
    util = require('util'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('The CoreProperties class', () => {
    let book;
    let coreProperties;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        coreProperties = book.coreProperties();
    });

    it('title and setTitle change the title', () => {
        shouldThrow(coreProperties.title, {});
        shouldThrow(coreProperties.setTitle, coreProperties, 1);

        expect(coreProperties.setTitle('Funky Title')).toBe(coreProperties);
        expect(coreProperties.title()).toBe('Funky Title');
    });
});
