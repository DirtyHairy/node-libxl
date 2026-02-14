const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
var xl = require('../lib/libxl'),
    util = require('util'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('The font class', function () {
    var book = new xl.Book(xl.BOOK_TYPE_XLS),
        font = book.addFont();

    it('The font constructor can not be called directly', function () {
        assert.throws(function () {
            new font.constructor();
        });
    });

    it('font.size reads font size', function () {
        shouldThrow(font.size, book);
        font.setSize(10);
        assert.strictEqual(font.size(), 10);
        font.setSize(20);
        assert.strictEqual(font.size(), 20);
    });

    it('font.setSize sets font size', function () {
        shouldThrow(font.setSize, font, '');
        shouldThrow(font.setSize, {}, 10);
        assert.strictEqual(font.setSize(10), font);
    });

    it('font.italic checks wether a font is italic', function () {
        shouldThrow(font.italic, book);
        font.setItalic(false);
        assert.strictEqual(font.italic(), false);
        font.setItalic(true);
        assert.strictEqual(font.italic(), true);
    });

    it('font.setItalic toggles a font italic', function () {
        shouldThrow(font.setItalic, font, 10);
        shouldThrow(font.setItalic, book, true);
        assert.strictEqual(font.setItalic(true), font);
    });

    it('font.strikeOut checks wether a font is striked out', function () {
        shouldThrow(font.strikeOut, book);
        font.setStrikeOut(false);
        assert.strictEqual(font.strikeOut(), false);
        font.setStrikeOut(true);
        assert.strictEqual(font.strikeOut(), true);
    });

    it('font.setStrikeOut toggles a font striked out', function () {
        shouldThrow(font.setStrikeOut, font, 10);
        shouldThrow(font.setStrikeOut, book, true);
        assert.strictEqual(font.setStrikeOut(true), font);
    });

    it('font.color gets font color', function () {
        shouldThrow(font.color, book);
        font.setColor(xl.COLOR_BLACK);
        assert.strictEqual(font.color(), xl.COLOR_BLACK);
        font.setColor(xl.COLOR_WHITE);
        assert.strictEqual(font.color(), xl.COLOR_WHITE);
    });

    it('font.setColor sets font color', function () {
        shouldThrow(font.setColor, font, true);
        shouldThrow(font.setColor, book, xl.COLOR_BLACK);
        assert.strictEqual(font.setColor(xl.COLOR_BLACK), font);
    });

    it('font.bold checks whether a font is bold', function () {
        shouldThrow(font.bold, book);
        font.setBold(false);
        assert.strictEqual(font.bold(), false);
        font.setBold(true);
        assert.strictEqual(font.bold(), true);
    });

    it('font.setBold toggles a font bold', function () {
        shouldThrow(font.setBold, font, 10);
        shouldThrow(font.setBold, book, true);
        assert.strictEqual(font.setBold(true), font);
    });

    it('font.script gets a fonts script style', function () {
        shouldThrow(font.script, book);
        font.setScript(xl.SCRIPT_NORMAL);
        assert.strictEqual(font.script(), xl.SCRIPT_NORMAL);
        font.setScript(xl.SCRIPT_SUPER);
        assert.strictEqual(font.script(), xl.SCRIPT_SUPER);
    });

    it('font.setScript sets a fonts script style', function () {
        shouldThrow(font.setScript, font, '');
        shouldThrow(font.setScript, book, xl.SCRIPT_NORMAL);
        assert.strictEqual(font.setScript(xl.SCRIPT_NORMAL), font);
    });

    it('font.underline gets a fonts underline mode', function () {
        shouldThrow(font.underline, book);
        font.setUnderline(xl.UNDERLINE_NONE);
        assert.strictEqual(font.underline(), xl.UNDERLINE_NONE);
        font.setUnderline(xl.UNDERLINE_SINGLE);
        assert.strictEqual(font.underline(), xl.UNDERLINE_SINGLE);
    });

    it('font.setUnderline sets a fonts underline mode', function () {
        shouldThrow(font.setUnderline, font, true);
        shouldThrow(font.setUnderline, book, xl.UNDERLINE_NONE);
        assert.strictEqual(font.setUnderline(xl.UNDERLINE_NONE), font);
    });

    it('font.name gets the font name', function () {
        shouldThrow(font.name, book);
        font.setName('arial');
        assert.strictEqual(font.name(), 'arial');
        font.setName('times');
        assert.strictEqual(font.name(), 'times');
    });

    it('font.setName sets the font name', function () {
        shouldThrow(font.setName, font, 10);
        shouldThrow(font.setName, book, 'arial');
        assert.strictEqual(font.setName('arial'), font);
    });
});
