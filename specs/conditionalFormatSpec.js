const { describe, it, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
var xl = require('../lib/libxl'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('ConditionalFormat', () => {
    let book, conditionalFormat;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        conditionalFormat = book.addConditionalFormat();
    });

    it('font gets the associated font', () => {
        shouldThrow(conditionalFormat.font, {});

        assert.ok(conditionalFormat.font() instanceof xl.Font);
    });

    it('numFormat and setNumFormat manage the number format', () => {
        shouldThrow(conditionalFormat.setNumFormat, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setNumFormat, {}, xl.NUMFORMAT_GENERAL);
        shouldThrow(conditionalFormat.numFormat, {});

        assert.strictEqual(conditionalFormat.setNumFormat(xl.NUMFORMAT_GENERAL), conditionalFormat);
        assert.strictEqual(conditionalFormat.numFormat(), xl.NUMFORMAT_GENERAL);
    });

    it('customNumFormat and setCustomNumFormat manage the number format', () => {
        shouldThrow(conditionalFormat.setCustomNumFormat, conditionalFormat, 10);
        shouldThrow(conditionalFormat.setCustomNumFormat, {}, 'abc');
        shouldThrow(conditionalFormat.customNumFormat, {});

        assert.strictEqual(conditionalFormat.setCustomNumFormat('abc'), conditionalFormat);
        assert.strictEqual(conditionalFormat.customNumFormat(), 'abc');
    });

    it('setBorder sets the border', () => {
        shouldThrow(conditionalFormat.setBorder, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setBorder, {}, xl.BORDERSTYLE_MEDIUM);
        shouldThrow(conditionalFormat.setBorder, {});

        assert.strictEqual(conditionalFormat.setBorder(xl.BORDERSTYLE_MEDIUM), conditionalFormat);
        assert.strictEqual(conditionalFormat.setBorder(), conditionalFormat);
    });

    it('setBorderColor sets the border color', () => {
        shouldThrow(conditionalFormat.setBorderColor, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setBorderColor, {}, xl.COLOR_RED);

        assert.strictEqual(conditionalFormat.setBorderColor(xl.COLOR_RED), conditionalFormat);
    });

    it('borderLeft and setBorderLeft manage left border', () => {
        shouldThrow(conditionalFormat.setBorderLeft, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setBorderLeft, {}, xl.BORDERSTYLE_MEDIUM);
        shouldThrow(conditionalFormat.borderLeft, {});

        assert.strictEqual(conditionalFormat.setBorderLeft(), conditionalFormat);
        assert.strictEqual(conditionalFormat.setBorderLeft(xl.BORDERSTYLE_MEDIUM), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderLeft(), xl.BORDERSTYLE_MEDIUM);
    });

    it('borderRight and setBorderRight manage right border', () => {
        shouldThrow(conditionalFormat.setBorderRight, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setBorderRight, {}, xl.BORDERSTYLE_MEDIUM);
        shouldThrow(conditionalFormat.borderRight, {});

        assert.strictEqual(conditionalFormat.setBorderRight(), conditionalFormat);
        assert.strictEqual(conditionalFormat.setBorderRight(xl.BORDERSTYLE_MEDIUM), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderRight(), xl.BORDERSTYLE_MEDIUM);
    });

    it('borderTop and setBorderTop manage top border', () => {
        shouldThrow(conditionalFormat.setBorderTop, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setBorderTop, {}, xl.BORDERSTYLE_MEDIUM);
        shouldThrow(conditionalFormat.borderTop, {});

        assert.strictEqual(conditionalFormat.setBorderTop(), conditionalFormat);
        assert.strictEqual(conditionalFormat.setBorderTop(xl.BORDERSTYLE_MEDIUM), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderTop(), xl.BORDERSTYLE_MEDIUM);
    });

    it('borderBottom and setBorderBottom manage bottom border', () => {
        shouldThrow(conditionalFormat.setBorderBottom, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setBorderBottom, {}, xl.BORDERSTYLE_MEDIUM);
        shouldThrow(conditionalFormat.borderBottom, {});

        assert.strictEqual(conditionalFormat.setBorderBottom(), conditionalFormat);
        assert.strictEqual(conditionalFormat.setBorderBottom(xl.BORDERSTYLE_MEDIUM), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderBottom(), xl.BORDERSTYLE_MEDIUM);
    });

    it('borderLeftColor and setBorderLeftColor manage left border color', () => {
        shouldThrow(conditionalFormat.borderLeftColor, {});
        shouldThrow(conditionalFormat.setBorderLeftColor, conditionalFormat, 'abc');

        assert.strictEqual(conditionalFormat.setBorderLeftColor(xl.COLOR_RED), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderLeftColor(), xl.COLOR_RED);
    });

    it('borderRightColor and setBorderRightColor manage right border color', () => {
        shouldThrow(conditionalFormat.borderRightColor, {});
        shouldThrow(conditionalFormat.setBorderRightColor, conditionalFormat, 'abc');

        assert.strictEqual(conditionalFormat.setBorderRightColor(xl.COLOR_GREEN), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderRightColor(), xl.COLOR_GREEN);
    });

    it('borderTopColor and setBorderTopColor manage top border color', () => {
        shouldThrow(conditionalFormat.borderTopColor, {});
        shouldThrow(conditionalFormat.setBorderTopColor, conditionalFormat, 'abc');

        assert.strictEqual(conditionalFormat.setBorderTopColor(xl.COLOR_BLUE), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderTopColor(), xl.COLOR_BLUE);
    });

    it('borderBottomColor and setBorderBottomColor manage bottom border color', () => {
        shouldThrow(conditionalFormat.borderBottomColor, {});
        shouldThrow(conditionalFormat.setBorderBottomColor, conditionalFormat, 'abc');

        assert.strictEqual(conditionalFormat.setBorderBottomColor(xl.COLOR_BLACK), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderBottomColor(), xl.COLOR_BLACK);
    });

    it('fillPattern and setFillPattern manage fill pattern', () => {
        shouldThrow(conditionalFormat.setFillPattern, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setFillPattern, {}, xl.FILLPATTERN_SOLID);
        shouldThrow(conditionalFormat.fillPattern, {});

        assert.strictEqual(conditionalFormat.setFillPattern(xl.FILLPATTERN_SOLID), conditionalFormat);
        assert.strictEqual(conditionalFormat.fillPattern(), xl.FILLPATTERN_SOLID);
    });

    it('patternForegroundColor and setPatternForegroundColor manage pattern foreground color', () => {
        shouldThrow(conditionalFormat.patternForegroundColor, {});
        shouldThrow(conditionalFormat.setPatternForegroundColor, conditionalFormat, 'abc');

        assert.strictEqual(conditionalFormat.setPatternForegroundColor(xl.COLOR_RED), conditionalFormat);
        assert.strictEqual(conditionalFormat.patternForegroundColor(), xl.COLOR_RED);
    });

    it('patternBackgroundColor and setPatternBackgroundColor manage pattern background color', () => {
        shouldThrow(conditionalFormat.patternBackgroundColor, {});
        shouldThrow(conditionalFormat.setPatternBackgroundColor, conditionalFormat, 'abc');

        assert.strictEqual(conditionalFormat.setPatternBackgroundColor(xl.COLOR_BLUE), conditionalFormat);
        assert.strictEqual(conditionalFormat.patternBackgroundColor(), xl.COLOR_BLUE);
    });
});
