var xl = require('../lib/libxl'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('RichString', () => {
    let book, richString, font;
    let anotherBook, anotherRichString, anotherFont;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        richString = book.addRichString();
        font = richString.addFont();

        anotherBook = new xl.Book(xl.BOOK_TYPE_XLSX);
        anotherRichString = anotherBook.addRichString();
        anotherFont = anotherRichString.addFont();
    });

    it('addFont adds a font', () => {
        shouldThrow(richString.addFont, richString, 1);
        shouldThrow(richString.addFont, richString, anotherFont);

        expect(richString.addFont()).toBeInstanceOf(xl.Font);
        expect(richString.addFont(font)).toBeInstanceOf(xl.Font);
    });

    it('addText and getText add end get text content', () => {
        shouldThrow(richString.addText, richString, 1);
        shouldThrow(richString.addText, {}, 'asd');
        shouldThrow(richString.addText, richString, anotherFont);

        shouldThrow(richString.getText, richString, 'a');
        shouldThrow(richString.getText, {}, 0);

        expect(richString.addText('asd')).toBe(richString);
        expect(richString.addText('asd', font)).toBe(richString);

        shouldThrow(richString.getText, richString, 'a');
        shouldThrow(richString.getText, {}, 0);

        const result = richString.getText(0);
        expect(result.text).toBe('asd');
        expect(result.font).toBeInstanceOf(xl.Font);
    });

    it('textSize returns the number of text elements"', () => {
        shouldThrow(richString.textSize, {});

        expect(richString.textSize()).toBe(0);

        richString.addText('asd');

        expect(richString.textSize()).toBe(1);
    });
});
