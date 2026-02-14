import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import * as xl from '../lib/libxl';

describe('The font class', () => {
    const book = new xl.Book(xl.BOOK_TYPE_XLS),
        font = book.addFont();

    it('The font constructor can not be called directly', () => {
        assert.throws(() => {
            new (font.constructor as any)();
        });
    });

    it('font.size reads font size', () => {
        assert.throws(() => (font.size as any).call(book));
        font.setSize(10);
        assert.strictEqual(font.size(), 10);
        font.setSize(20);
        assert.strictEqual(font.size(), 20);
    });

    it('font.setSize sets font size', () => {
        assert.throws(() => (font.setSize as any).call(font, ''));
        assert.throws(() => (font.setSize as any).call({}, 10));
        assert.strictEqual(font.setSize(10), font);
    });

    it('font.italic checks wether a font is italic', () => {
        assert.throws(() => (font.italic as any).call(book));
        font.setItalic(false);
        assert.strictEqual(font.italic(), false);
        font.setItalic(true);
        assert.strictEqual(font.italic(), true);
    });

    it('font.setItalic toggles a font italic', () => {
        assert.throws(() => (font.setItalic as any).call(font, 10));
        assert.throws(() => (font.setItalic as any).call(book, true));
        assert.strictEqual(font.setItalic(true), font);
    });

    it('font.strikeOut checks wether a font is striked out', () => {
        assert.throws(() => (font.strikeOut as any).call(book));
        font.setStrikeOut(false);
        assert.strictEqual(font.strikeOut(), false);
        font.setStrikeOut(true);
        assert.strictEqual(font.strikeOut(), true);
    });

    it('font.setStrikeOut toggles a font striked out', () => {
        assert.throws(() => (font.setStrikeOut as any).call(font, 10));
        assert.throws(() => (font.setStrikeOut as any).call(book, true));
        assert.strictEqual(font.setStrikeOut(true), font);
    });

    it('font.color gets font color', () => {
        assert.throws(() => (font.color as any).call(book));
        font.setColor(xl.COLOR_BLACK);
        assert.strictEqual(font.color(), xl.COLOR_BLACK);
        font.setColor(xl.COLOR_WHITE);
        assert.strictEqual(font.color(), xl.COLOR_WHITE);
    });

    it('font.setColor sets font color', () => {
        assert.throws(() => (font.setColor as any).call(font, true));
        assert.throws(() => (font.setColor as any).call(book, xl.COLOR_BLACK));
        assert.strictEqual(font.setColor(xl.COLOR_BLACK), font);
    });

    it('font.bold checks whether a font is bold', () => {
        assert.throws(() => (font.bold as any).call(book));
        font.setBold(false);
        assert.strictEqual(font.bold(), false);
        font.setBold(true);
        assert.strictEqual(font.bold(), true);
    });

    it('font.setBold toggles a font bold', () => {
        assert.throws(() => (font.setBold as any).call(font, 10));
        assert.throws(() => (font.setBold as any).call(book, true));
        assert.strictEqual(font.setBold(true), font);
    });

    it('font.script gets a fonts script style', () => {
        assert.throws(() => (font.script as any).call(book));
        font.setScript(xl.SCRIPT_NORMAL);
        assert.strictEqual(font.script(), xl.SCRIPT_NORMAL);
        font.setScript(xl.SCRIPT_SUPER);
        assert.strictEqual(font.script(), xl.SCRIPT_SUPER);
    });

    it('font.setScript sets a fonts script style', () => {
        assert.throws(() => (font.setScript as any).call(font, ''));
        assert.throws(() => (font.setScript as any).call(book, xl.SCRIPT_NORMAL));
        assert.strictEqual(font.setScript(xl.SCRIPT_NORMAL), font);
    });

    it('font.underline gets a fonts underline mode', () => {
        assert.throws(() => (font.underline as any).call(book));
        font.setUnderline(xl.UNDERLINE_NONE);
        assert.strictEqual(font.underline(), xl.UNDERLINE_NONE);
        font.setUnderline(xl.UNDERLINE_SINGLE);
        assert.strictEqual(font.underline(), xl.UNDERLINE_SINGLE);
    });

    it('font.setUnderline sets a fonts underline mode', () => {
        assert.throws(() => (font.setUnderline as any).call(font, true));
        assert.throws(() => (font.setUnderline as any).call(book, xl.UNDERLINE_NONE));
        assert.strictEqual(font.setUnderline(xl.UNDERLINE_NONE), font);
    });

    it('font.name gets the font name', () => {
        assert.throws(() => (font.name as any).call(book));
        font.setName('arial');
        assert.strictEqual(font.name(), 'arial');
        font.setName('times');
        assert.strictEqual(font.name(), 'times');
    });

    it('font.setName sets the font name', () => {
        assert.throws(() => (font.setName as any).call(font, 10));
        assert.throws(() => (font.setName as any).call(book, 'arial'));
        assert.strictEqual(font.setName('arial'), font);
    });
});
