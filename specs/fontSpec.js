var xl = require('../lib/libxl'),
    util = require('util'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('The font class', function() {

    var book = new xl.Book(xl.BOOK_TYPE_XLS),
        font = book.addFont();

    it('The font constructor can not be called directly', function() {
        expect(function() {new font.constructor();}).toThrow();
    });

    it('font.size reads font size', function() {
        shouldThrow(font.size, book);
        font.setSize(10);
        expect(font.size()).toBe(10);
        font.setSize(20);
        expect(font.size()).toBe(20);
    });

    it('font.setSize sets font size', function() {
        shouldThrow(font.setSize, font, '');
        shouldThrow(font.setSize, {}, 10);
        expect(font.setSize(10)).toBe(font);
    });

    it('font.italic checks wether a font is italic', function() {
        shouldThrow(font.italic, book);
        font.setItalic(false);
        expect(font.italic()).toBe(false);
        font.setItalic(true);
        expect(font.italic()).toBe(true);
    });

    it('font.setItalic toggles a font italic', function() {
        shouldThrow(font.setItalic, font, 10);
        shouldThrow(font.setItalic, book, true);
        expect(font.setItalic(true)).toBe(font);
    });

    it('font.strikeOut checks wether a font is striked out', function() {
        shouldThrow(font.strikeOut, book);
        font.setStrikeOut(false);
        expect(font.strikeOut()).toBe(false);
        font.setStrikeOut(true);
        expect(font.strikeOut()).toBe(true);
    });

    it('font.setStrikeOut toggles a font striked out', function() {
        shouldThrow(font.setStrikeOut, font, 10);
        shouldThrow(font.setStrikeOut, book, true);
        expect(font.setStrikeOut(true)).toBe(font);
    });
   

    it('font.color gets font color', function() {
        shouldThrow(font.color, book);
        font.setColor(xl.COLOR_BLACK);
        expect(font.color()).toBe(xl.COLOR_BLACK);
        font.setColor(xl.COLOR_WHITE);
        expect(font.color()).toBe(xl.COLOR_WHITE);
    });

    it('font.setColor sets font color', function() {
        shouldThrow(font.setColor, font, true);
        shouldThrow(font.setColor, book, xl.COLOR_BLACK);
        expect(font.setColor(xl.COLOR_BLACK)).toBe(font);
    });

    it('font.bold checks whether a font is bold', function() {
        shouldThrow(font.bold, book);
        font.setBold(false);
        expect(font.bold()).toBe(false);
        font.setBold(true);
        expect(font.bold()).toBe(true);
    });

    it('font.setBold toggles a font bold', function() {
        shouldThrow(font.setBold, font, 10);
        shouldThrow(font.setBold, book, true);
        expect(font.setBold(true)).toBe(font);
    });

    it('font.script gets a fonts script style', function() {
        shouldThrow(font.script, book);
        font.setScript(xl.SCRIPT_NORMAL);
        expect(font.script()).toBe(xl.SCRIPT_NORMAL);
        font.setScript(xl.SCRIPT_SUPER);
        expect(font.script()).toBe(xl.SCRIPT_SUPER);
    });

    it('font.setScript sets a fonts script style', function() {
        shouldThrow(font.setScript, font, '');
        shouldThrow(font.setScript, book, xl.SCRIPT_NORMAL);
        expect(font.setScript(xl.SCRIPT_NORMAL)).toBe(font);
    });

    it('font.underline gets a fonts underline mode', function() {
        shouldThrow(font.underline, book);
        font.setUnderline(xl.UNDERLINE_NONE);
        expect(font.underline()).toBe(xl.UNDERLINE_NONE);
        font.setUnderline(xl.UNDERLINE_SINGLE);
        expect(font.underline()).toBe(xl.UNDERLINE_SINGLE);
    });

    it('font.setUnderline sets a fonts underline mode', function() {
        shouldThrow(font.setUnderline, font, true);
        shouldThrow(font.setUnderline, book, xl.UNDERLINE_NONE);
        expect(font.setUnderline(xl.UNDERLINE_NONE)).toBe(font);
    });

    it('font.name gets the font name', function() {
        shouldThrow(font.name, book);
        font.setName('arial');
        expect(font.name()).toBe('arial');
        font.setName('times');
        expect(font.name()).toBe('times');
    });

    it('font.setName sets the font name', function() {
        shouldThrow(font.setName, font, 10);
        shouldThrow(font.setName, book, 'arial');
        expect(font.setName('arial')).toBe(font);
    });
});

