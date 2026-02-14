const { describe, it, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
var xl = require('../lib/libxl'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('ConditionalFormatting', () => {
    let book, sheet, conditionalFormat, conditionalFormatting;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        sheet = book.addSheet('Sheet1');
        conditionalFormatting = sheet.addConditionalFormatting(0, 0, 1, 1);
        conditionalFormat = book.addConditionalFormat();
    });

    it('addRange adds a range', () => {
        shouldThrow(conditionalFormatting.addRange, conditionalFormatting, 1, 2, 3, 'a');
        shouldThrow(conditionalFormatting.addRange, {}, 1, 2, 3, 4);

        assert.strictEqual(conditionalFormatting.addRange(1, 2, 3, 4), conditionalFormatting);
    });

    it('addRule adds a rule', () => {
        shouldThrow(conditionalFormatting.addRule, conditionalFormatting, xl.CFORMAT_BEGINWITH, 'a');
        shouldThrow(conditionalFormatting.addRule, {}, xl.CFORMAT_BEGINWITH, conditionalFormat);

        assert.strictEqual(
            conditionalFormatting.addRule(xl.CFORMAT_BEGINWITH, conditionalFormat),
            conditionalFormatting,
        );
        assert.strictEqual(
            conditionalFormatting.addRule(xl.CFORMAT_BEGINWITH, conditionalFormat, 'a'),
            conditionalFormatting,
        );
        assert.strictEqual(
            conditionalFormatting.addRule(xl.CFORMAT_BEGINWITH, conditionalFormat, 'a', false),
            conditionalFormatting,
        );
    });

    it('addTopRule adds a top rule', () => {
        shouldThrow(conditionalFormatting.addTopRule, conditionalFormatting, 'invalid', 10);
        shouldThrow(conditionalFormatting.addTopRule, {}, conditionalFormat, 10);

        assert.strictEqual(conditionalFormatting.addTopRule(conditionalFormat, 10), conditionalFormatting);
        assert.strictEqual(conditionalFormatting.addTopRule(conditionalFormat, 10, true), conditionalFormatting);
        assert.strictEqual(conditionalFormatting.addTopRule(conditionalFormat, 10, true, true), conditionalFormatting);
        assert.strictEqual(
            conditionalFormatting.addTopRule(conditionalFormat, 10, true, true, true),
            conditionalFormatting,
        );
    });

    it('addOpNumRule adds a numeric operator rule', () => {
        shouldThrow(conditionalFormatting.addOpNumRule, conditionalFormatting, xl.CFOPERATOR_LESSTHAN, 'invalid', 10);
        shouldThrow(conditionalFormatting.addOpNumRule, {}, xl.CFOPERATOR_LESSTHAN, conditionalFormat, 10);

        assert.strictEqual(
            conditionalFormatting.addOpNumRule(xl.CFOPERATOR_LESSTHAN, conditionalFormat, 10),
            conditionalFormatting,
        );
        assert.strictEqual(
            conditionalFormatting.addOpNumRule(xl.CFOPERATOR_LESSTHAN, conditionalFormat, 10, 20),
            conditionalFormatting,
        );
        assert.strictEqual(
            conditionalFormatting.addOpNumRule(xl.CFOPERATOR_LESSTHAN, conditionalFormat, 10, 0, true),
            conditionalFormatting,
        );
    });

    it('addOpStrRule adds a string operator rule', () => {
        shouldThrow(
            conditionalFormatting.addOpStrRule,
            conditionalFormatting,
            xl.CFOPERATOR_LESSTHAN,
            'invalid',
            'test',
        );
        shouldThrow(conditionalFormatting.addOpStrRule, {}, xl.CFOPERATOR_LESSTHAN, conditionalFormat, 'test');

        assert.strictEqual(
            conditionalFormatting.addOpStrRule(xl.CFOPERATOR_LESSTHAN, conditionalFormat, 'test'),
            conditionalFormatting,
        );
        assert.strictEqual(
            conditionalFormatting.addOpStrRule(xl.CFOPERATOR_LESSTHAN, conditionalFormat, 'a', 'z'),
            conditionalFormatting,
        );
        assert.strictEqual(
            conditionalFormatting.addOpStrRule(xl.CFOPERATOR_LESSTHAN, conditionalFormat, 'test', '', true),
            conditionalFormatting,
        );
    });

    it('addAboveAverageRule adds an above average rule', () => {
        shouldThrow(conditionalFormatting.addAboveAverageRule, conditionalFormatting, 'invalid');
        shouldThrow(conditionalFormatting.addAboveAverageRule, {}, conditionalFormat);

        assert.strictEqual(conditionalFormatting.addAboveAverageRule(conditionalFormat), conditionalFormatting);
        assert.strictEqual(conditionalFormatting.addAboveAverageRule(conditionalFormat, false), conditionalFormatting);
        assert.strictEqual(
            conditionalFormatting.addAboveAverageRule(conditionalFormat, true, true),
            conditionalFormatting,
        );
        assert.strictEqual(
            conditionalFormatting.addAboveAverageRule(conditionalFormat, true, false, 1),
            conditionalFormatting,
        );
        assert.strictEqual(
            conditionalFormatting.addAboveAverageRule(conditionalFormat, true, false, 1, true),
            conditionalFormatting,
        );
    });

    it('addTimePeriodRule adds a time period rule', () => {
        shouldThrow(conditionalFormatting.addTimePeriodRule, conditionalFormatting, 'invalid', xl.CFTP_LAST7DAYS);
        shouldThrow(conditionalFormatting.addTimePeriodRule, {}, conditionalFormat, xl.CFTP_LAST7DAYS);

        assert.strictEqual(
            conditionalFormatting.addTimePeriodRule(conditionalFormat, xl.CFTP_LAST7DAYS),
            conditionalFormatting,
        );
        assert.strictEqual(
            conditionalFormatting.addTimePeriodRule(conditionalFormat, xl.CFTP_LAST7DAYS, true),
            conditionalFormatting,
        );
    });

    it('add2ColorScaleRule adds a 2-color scale rule', () => {
        shouldThrow(conditionalFormatting.add2ColorScaleRule, conditionalFormatting, 'invalid', xl.COLOR_RED);
        shouldThrow(conditionalFormatting.add2ColorScaleRule, {}, xl.COLOR_GREEN, xl.COLOR_RED);

        assert.strictEqual(
            conditionalFormatting.add2ColorScaleRule(xl.COLOR_GREEN, xl.COLOR_RED),
            conditionalFormatting,
        );
        assert.strictEqual(
            conditionalFormatting.add2ColorScaleRule(
                xl.COLOR_GREEN,
                xl.COLOR_RED,
                xl.CFVO_MIN,
                0,
                xl.CFVO_MAX,
                100,
                true,
            ),
            conditionalFormatting,
        );
    });

    it('add2ColorScaleFormulaRule adds a 2-color scale formula rule', () => {
        shouldThrow(conditionalFormatting.add2ColorScaleFormulaRule, conditionalFormatting, 'invalid', xl.COLOR_RED);
        shouldThrow(conditionalFormatting.add2ColorScaleFormulaRule, {}, xl.COLOR_GREEN, xl.COLOR_RED);

        assert.strictEqual(
            conditionalFormatting.add2ColorScaleFormulaRule(xl.COLOR_GREEN, xl.COLOR_RED),
            conditionalFormatting,
        );
        assert.strictEqual(
            conditionalFormatting.add2ColorScaleFormulaRule(
                xl.COLOR_GREEN,
                xl.COLOR_RED,
                xl.CFVO_FORMULA,
                'A1',
                xl.CFVO_FORMULA,
                'B1',
                true,
            ),
            conditionalFormatting,
        );
    });

    it('add3ColorScaleRule adds a 3-color scale rule', () => {
        shouldThrow(
            conditionalFormatting.add3ColorScaleRule,
            conditionalFormatting,
            'invalid',
            xl.COLOR_YELLOW,
            xl.COLOR_RED,
        );
        shouldThrow(conditionalFormatting.add3ColorScaleRule, {}, xl.COLOR_GREEN, xl.COLOR_YELLOW, xl.COLOR_RED);

        assert.strictEqual(
            conditionalFormatting.add3ColorScaleRule(xl.COLOR_GREEN, xl.COLOR_YELLOW, xl.COLOR_RED),
            conditionalFormatting,
        );
        assert.strictEqual(
            conditionalFormatting.add3ColorScaleRule(
                xl.COLOR_GREEN,
                xl.COLOR_YELLOW,
                xl.COLOR_RED,
                xl.CFVO_MIN,
                0,
                xl.CFVO_PERCENTILE,
                50,
                xl.CFVO_MAX,
                100,
                true,
            ),
            conditionalFormatting,
        );
    });

    it('add3ColorScaleFormulaRule adds a 3-color scale formula rule', () => {
        shouldThrow(
            conditionalFormatting.add3ColorScaleFormulaRule,
            conditionalFormatting,
            'invalid',
            xl.COLOR_YELLOW,
            xl.COLOR_RED,
        );
        shouldThrow(conditionalFormatting.add3ColorScaleFormulaRule, {}, xl.COLOR_GREEN, xl.COLOR_YELLOW, xl.COLOR_RED);

        assert.strictEqual(
            conditionalFormatting.add3ColorScaleFormulaRule(xl.COLOR_GREEN, xl.COLOR_YELLOW, xl.COLOR_RED),
            conditionalFormatting,
        );
        assert.strictEqual(
            conditionalFormatting.add3ColorScaleFormulaRule(
                xl.COLOR_GREEN,
                xl.COLOR_YELLOW,
                xl.COLOR_RED,
                xl.CFVO_FORMULA,
                'A1',
                xl.CFVO_FORMULA,
                'B1',
                xl.CFVO_FORMULA,
                'C1',
                true,
            ),
            conditionalFormatting,
        );
    });
});
