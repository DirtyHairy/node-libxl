const { describe, it, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
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

        assert.ok(richString.addFont() instanceof xl.Font);
        assert.ok(richString.addFont(font) instanceof xl.Font);
    });

    it('addText and getText add end get text content', () => {
        shouldThrow(richString.addText, richString, 1);
        shouldThrow(richString.addText, {}, 'asd');
        shouldThrow(richString.addText, richString, anotherFont);

        shouldThrow(richString.getText, richString, 'a');
        shouldThrow(richString.getText, {}, 0);

        assert.strictEqual(richString.addText('asd'), richString);
        assert.strictEqual(richString.addText('asd', font), richString);

        shouldThrow(richString.getText, richString, 'a');
        shouldThrow(richString.getText, {}, 0);

        const result = richString.getText(0);
        assert.strictEqual(result.text, 'asd');
        assert.ok(result.font instanceof xl.Font);
    });

    it('textSize returns the number of text elements"', () => {
        shouldThrow(richString.textSize, {});

        assert.strictEqual(richString.textSize(), 0);

        richString.addText('asd');

        assert.strictEqual(richString.textSize(), 1);
    });
});
