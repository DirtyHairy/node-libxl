import { describe, it, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import * as xl from '../lib/libxl.js';

describe('ConditionalFormatting', () => {
    let book: xl.Book,
        sheet: xl.Sheet,
        conditionalFormat: xl.ConditionalFormat,
        conditionalFormatting: xl.ConditionalFormatting;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        sheet = book.addSheet('Sheet1');
        conditionalFormatting = sheet.addConditionalFormatting(0, 0, 1, 1);
        conditionalFormat = book.addConditionalFormat();
    });

    it('addRange adds a range', () => {
        assert.throws(() => (conditionalFormatting.addRange as any).call(conditionalFormatting, 1, 2, 3, 'a'));
        assert.throws(() => (conditionalFormatting.addRange as any).call({}, 1, 2, 3, 4));

        assert.strictEqual(conditionalFormatting.addRange(1, 2, 3, 4), conditionalFormatting);
    });

    it('addRule adds a rule', () => {
        assert.throws(() =>
            (conditionalFormatting.addRule as any).call(conditionalFormatting, xl.CFORMAT_BEGINWITH, 'a'),
        );
        assert.throws(() => (conditionalFormatting.addRule as any).call({}, xl.CFORMAT_BEGINWITH, conditionalFormat));

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
        assert.throws(() => (conditionalFormatting.addTopRule as any).call(conditionalFormatting, 'invalid', 10));
        assert.throws(() => (conditionalFormatting.addTopRule as any).call({}, conditionalFormat, 10));

        assert.strictEqual(conditionalFormatting.addTopRule(conditionalFormat, 10), conditionalFormatting);
        assert.strictEqual(conditionalFormatting.addTopRule(conditionalFormat, 10, true), conditionalFormatting);
        assert.strictEqual(conditionalFormatting.addTopRule(conditionalFormat, 10, true, true), conditionalFormatting);
        assert.strictEqual(
            conditionalFormatting.addTopRule(conditionalFormat, 10, true, true, true),
            conditionalFormatting,
        );
    });

    it('addOpNumRule adds a numeric operator rule', () => {
        assert.throws(() =>
            (conditionalFormatting.addOpNumRule as any).call(
                conditionalFormatting,
                xl.CFOPERATOR_LESSTHAN,
                'invalid',
                10,
            ),
        );
        assert.throws(() =>
            (conditionalFormatting.addOpNumRule as any).call({}, xl.CFOPERATOR_LESSTHAN, conditionalFormat, 10),
        );

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
        assert.throws(() =>
            (conditionalFormatting.addOpStrRule as any).call(
                conditionalFormatting,
                xl.CFOPERATOR_LESSTHAN,
                'invalid',
                'test',
            ),
        );
        assert.throws(() =>
            (conditionalFormatting.addOpStrRule as any).call({}, xl.CFOPERATOR_LESSTHAN, conditionalFormat, 'test'),
        );

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
        assert.throws(() => (conditionalFormatting.addAboveAverageRule as any).call(conditionalFormatting, 'invalid'));
        assert.throws(() => (conditionalFormatting.addAboveAverageRule as any).call({}, conditionalFormat));

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
        assert.throws(() =>
            (conditionalFormatting.addTimePeriodRule as any).call(conditionalFormatting, 'invalid', xl.CFTP_LAST7DAYS),
        );
        assert.throws(() =>
            (conditionalFormatting.addTimePeriodRule as any).call({}, conditionalFormat, xl.CFTP_LAST7DAYS),
        );

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
        assert.throws(() =>
            (conditionalFormatting.add2ColorScaleRule as any).call(conditionalFormatting, 'invalid', xl.COLOR_RED),
        );
        assert.throws(() => (conditionalFormatting.add2ColorScaleRule as any).call({}, xl.COLOR_GREEN, xl.COLOR_RED));

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
        assert.throws(() =>
            (conditionalFormatting.add2ColorScaleFormulaRule as any).call(
                conditionalFormatting,
                'invalid',
                xl.COLOR_RED,
            ),
        );
        assert.throws(() =>
            (conditionalFormatting.add2ColorScaleFormulaRule as any).call({}, xl.COLOR_GREEN, xl.COLOR_RED),
        );

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
        assert.throws(() =>
            (conditionalFormatting.add3ColorScaleRule as any).call(
                conditionalFormatting,
                'invalid',
                xl.COLOR_YELLOW,
                xl.COLOR_RED,
            ),
        );
        assert.throws(() =>
            (conditionalFormatting.add3ColorScaleRule as any).call({}, xl.COLOR_GREEN, xl.COLOR_YELLOW, xl.COLOR_RED),
        );

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
        assert.throws(() =>
            (conditionalFormatting.add3ColorScaleFormulaRule as any).call(
                conditionalFormatting,
                'invalid',
                xl.COLOR_YELLOW,
                xl.COLOR_RED,
            ),
        );
        assert.throws(() =>
            (conditionalFormatting.add3ColorScaleFormulaRule as any).call(
                {},
                xl.COLOR_GREEN,
                xl.COLOR_YELLOW,
                xl.COLOR_RED,
            ),
        );

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
