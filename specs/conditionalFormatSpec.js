import { describe, it, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import xl from '../lib/libxl.js';

describe('ConditionalFormat', () => {
    let book, conditionalFormat;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        conditionalFormat = book.addConditionalFormat();
    });

    it('font gets the associated font', () => {
        assert.throws(() => conditionalFormat.font.call({}));

        assert.ok(conditionalFormat.font() instanceof xl.Font);
    });

    it('numFormat and setNumFormat manage the number format', () => {
        assert.throws(() => conditionalFormat.setNumFormat.call(conditionalFormat, 'abc'));
        assert.throws(() => conditionalFormat.setNumFormat.call({}, xl.NUMFORMAT_GENERAL));
        assert.throws(() => conditionalFormat.numFormat.call({}));

        assert.strictEqual(conditionalFormat.setNumFormat(xl.NUMFORMAT_GENERAL), conditionalFormat);
        assert.strictEqual(conditionalFormat.numFormat(), xl.NUMFORMAT_GENERAL);
    });

    it('customNumFormat and setCustomNumFormat manage the number format', () => {
        assert.throws(() => conditionalFormat.setCustomNumFormat.call(conditionalFormat, 10));
        assert.throws(() => conditionalFormat.setCustomNumFormat.call({}, 'abc'));
        assert.throws(() => conditionalFormat.customNumFormat.call({}));

        assert.strictEqual(conditionalFormat.setCustomNumFormat('abc'), conditionalFormat);
        assert.strictEqual(conditionalFormat.customNumFormat(), 'abc');
    });

    it('setBorder sets the border', () => {
        assert.throws(() => conditionalFormat.setBorder.call(conditionalFormat, 'abc'));
        assert.throws(() => conditionalFormat.setBorder.call({}, xl.BORDERSTYLE_MEDIUM));
        assert.throws(() => conditionalFormat.setBorder.call({}));

        assert.strictEqual(conditionalFormat.setBorder(xl.BORDERSTYLE_MEDIUM), conditionalFormat);
        assert.strictEqual(conditionalFormat.setBorder(), conditionalFormat);
    });

    it('setBorderColor sets the border color', () => {
        assert.throws(() => conditionalFormat.setBorderColor.call(conditionalFormat, 'abc'));
        assert.throws(() => conditionalFormat.setBorderColor.call({}, xl.COLOR_RED));

        assert.strictEqual(conditionalFormat.setBorderColor(xl.COLOR_RED), conditionalFormat);
    });

    it('borderLeft and setBorderLeft manage left border', () => {
        assert.throws(() => conditionalFormat.setBorderLeft.call(conditionalFormat, 'abc'));
        assert.throws(() => conditionalFormat.setBorderLeft.call({}, xl.BORDERSTYLE_MEDIUM));
        assert.throws(() => conditionalFormat.borderLeft.call({}));

        assert.strictEqual(conditionalFormat.setBorderLeft(), conditionalFormat);
        assert.strictEqual(conditionalFormat.setBorderLeft(xl.BORDERSTYLE_MEDIUM), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderLeft(), xl.BORDERSTYLE_MEDIUM);
    });

    it('borderRight and setBorderRight manage right border', () => {
        assert.throws(() => conditionalFormat.setBorderRight.call(conditionalFormat, 'abc'));
        assert.throws(() => conditionalFormat.setBorderRight.call({}, xl.BORDERSTYLE_MEDIUM));
        assert.throws(() => conditionalFormat.borderRight.call({}));

        assert.strictEqual(conditionalFormat.setBorderRight(), conditionalFormat);
        assert.strictEqual(conditionalFormat.setBorderRight(xl.BORDERSTYLE_MEDIUM), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderRight(), xl.BORDERSTYLE_MEDIUM);
    });

    it('borderTop and setBorderTop manage top border', () => {
        assert.throws(() => conditionalFormat.setBorderTop.call(conditionalFormat, 'abc'));
        assert.throws(() => conditionalFormat.setBorderTop.call({}, xl.BORDERSTYLE_MEDIUM));
        assert.throws(() => conditionalFormat.borderTop.call({}));

        assert.strictEqual(conditionalFormat.setBorderTop(), conditionalFormat);
        assert.strictEqual(conditionalFormat.setBorderTop(xl.BORDERSTYLE_MEDIUM), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderTop(), xl.BORDERSTYLE_MEDIUM);
    });

    it('borderBottom and setBorderBottom manage bottom border', () => {
        assert.throws(() => conditionalFormat.setBorderBottom.call(conditionalFormat, 'abc'));
        assert.throws(() => conditionalFormat.setBorderBottom.call({}, xl.BORDERSTYLE_MEDIUM));
        assert.throws(() => conditionalFormat.borderBottom.call({}));

        assert.strictEqual(conditionalFormat.setBorderBottom(), conditionalFormat);
        assert.strictEqual(conditionalFormat.setBorderBottom(xl.BORDERSTYLE_MEDIUM), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderBottom(), xl.BORDERSTYLE_MEDIUM);
    });

    it('borderLeftColor and setBorderLeftColor manage left border color', () => {
        assert.throws(() => conditionalFormat.borderLeftColor.call({}));
        assert.throws(() => conditionalFormat.setBorderLeftColor.call(conditionalFormat, 'abc'));

        assert.strictEqual(conditionalFormat.setBorderLeftColor(xl.COLOR_RED), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderLeftColor(), xl.COLOR_RED);
    });

    it('borderRightColor and setBorderRightColor manage right border color', () => {
        assert.throws(() => conditionalFormat.borderRightColor.call({}));
        assert.throws(() => conditionalFormat.setBorderRightColor.call(conditionalFormat, 'abc'));

        assert.strictEqual(conditionalFormat.setBorderRightColor(xl.COLOR_GREEN), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderRightColor(), xl.COLOR_GREEN);
    });

    it('borderTopColor and setBorderTopColor manage top border color', () => {
        assert.throws(() => conditionalFormat.borderTopColor.call({}));
        assert.throws(() => conditionalFormat.setBorderTopColor.call(conditionalFormat, 'abc'));

        assert.strictEqual(conditionalFormat.setBorderTopColor(xl.COLOR_BLUE), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderTopColor(), xl.COLOR_BLUE);
    });

    it('borderBottomColor and setBorderBottomColor manage bottom border color', () => {
        assert.throws(() => conditionalFormat.borderBottomColor.call({}));
        assert.throws(() => conditionalFormat.setBorderBottomColor.call(conditionalFormat, 'abc'));

        assert.strictEqual(conditionalFormat.setBorderBottomColor(xl.COLOR_BLACK), conditionalFormat);
        assert.strictEqual(conditionalFormat.borderBottomColor(), xl.COLOR_BLACK);
    });

    it('fillPattern and setFillPattern manage fill pattern', () => {
        assert.throws(() => conditionalFormat.setFillPattern.call(conditionalFormat, 'abc'));
        assert.throws(() => conditionalFormat.setFillPattern.call({}, xl.FILLPATTERN_SOLID));
        assert.throws(() => conditionalFormat.fillPattern.call({}));

        assert.strictEqual(conditionalFormat.setFillPattern(xl.FILLPATTERN_SOLID), conditionalFormat);
        assert.strictEqual(conditionalFormat.fillPattern(), xl.FILLPATTERN_SOLID);
    });

    it('patternForegroundColor and setPatternForegroundColor manage pattern foreground color', () => {
        assert.throws(() => conditionalFormat.patternForegroundColor.call({}));
        assert.throws(() => conditionalFormat.setPatternForegroundColor.call(conditionalFormat, 'abc'));

        assert.strictEqual(conditionalFormat.setPatternForegroundColor(xl.COLOR_RED), conditionalFormat);
        assert.strictEqual(conditionalFormat.patternForegroundColor(), xl.COLOR_RED);
    });

    it('patternBackgroundColor and setPatternBackgroundColor manage pattern background color', () => {
        assert.throws(() => conditionalFormat.patternBackgroundColor.call({}));
        assert.throws(() => conditionalFormat.setPatternBackgroundColor.call(conditionalFormat, 'abc'));

        assert.strictEqual(conditionalFormat.setPatternBackgroundColor(xl.COLOR_BLUE), conditionalFormat);
        assert.strictEqual(conditionalFormat.patternBackgroundColor(), xl.COLOR_BLUE);
    });
});
