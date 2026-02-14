import { describe, it, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import xl from '../lib/libxl.js';

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
        assert.throws(() => richString.addFont.call(richString, 1));
        assert.throws(() => richString.addFont.call(richString, anotherFont));

        assert.ok(richString.addFont() instanceof xl.Font);
        assert.ok(richString.addFont(font) instanceof xl.Font);
    });

    it('addText and getText add end get text content', () => {
        assert.throws(() => richString.addText.call(richString, 1));
        assert.throws(() => richString.addText.call({}, 'asd'));
        assert.throws(() => richString.addText.call(richString, anotherFont));

        assert.throws(() => richString.getText.call(richString, 'a'));
        assert.throws(() => richString.getText.call({}, 0));

        assert.strictEqual(richString.addText('asd'), richString);
        assert.strictEqual(richString.addText('asd', font), richString);

        assert.throws(() => richString.getText.call(richString, 'a'));
        assert.throws(() => richString.getText.call({}, 0));

        const result = richString.getText(0);
        assert.strictEqual(result.text, 'asd');
        assert.ok(result.font instanceof xl.Font);
    });

    it('textSize returns the number of text elements"', () => {
        assert.throws(() => richString.textSize.call({}));

        assert.strictEqual(richString.textSize(), 0);

        richString.addText('asd');

        assert.strictEqual(richString.textSize(), 1);
    });
});
