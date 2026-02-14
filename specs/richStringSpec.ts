import { describe, it, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import * as xl from '../lib/libxl';

describe('RichString', () => {
    let book: xl.Book, richString: xl.RichString, font: xl.Font;
    let anotherBook: xl.Book, anotherRichString: xl.RichString, anotherFont: xl.Font;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        richString = book.addRichString();
        font = richString.addFont();

        anotherBook = new xl.Book(xl.BOOK_TYPE_XLSX);
        anotherRichString = anotherBook.addRichString();
        anotherFont = anotherRichString.addFont();
    });

    it('addFont adds a font', () => {
        assert.throws(() => (richString.addFont as any).call(richString, 1));
        assert.throws(() => (richString.addFont as any).call(richString, anotherFont));

        assert.ok(richString.addFont() instanceof xl.Font);
        assert.ok(richString.addFont(font) instanceof xl.Font);
    });

    it('addText and getText add end get text content', () => {
        assert.throws(() => (richString.addText as any).call(richString, 1));
        assert.throws(() => (richString.addText as any).call({}, 'asd'));
        assert.throws(() => (richString.addText as any).call(richString, anotherFont));

        assert.throws(() => (richString.getText as any).call(richString, 'a'));
        assert.throws(() => (richString.getText as any).call({}, 0));

        assert.strictEqual(richString.addText('asd'), richString);
        assert.strictEqual(richString.addText('asd', font), richString);

        assert.throws(() => (richString.getText as any).call(richString, 'a'));
        assert.throws(() => (richString.getText as any).call({}, 0));

        const result = richString.getText(0);
        assert.strictEqual(result.text, 'asd');
        assert.ok(result.font instanceof xl.Font);
    });

    it('textSize returns the number of text elements"', () => {
        assert.throws(() => (richString.textSize as any).call({}));

        assert.strictEqual(richString.textSize(), 0);

        richString.addText('asd');

        assert.strictEqual(richString.textSize(), 1);
    });
});
