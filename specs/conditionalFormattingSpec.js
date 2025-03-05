var xl = require('../lib/libxl'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('ConditionalFormatting', () => {
    let book, sheet, conditionalFormat, conditionalFormatting;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        sheet = book.addSheet('Sheet1');
        conditionalFormatting = sheet.addConditionalFormatting();
        conditionalFormat = book.addConditionalFormat();
    });

    it('addRange adds a range', () => {
        shouldThrow(conditionalFormatting.addRange, conditionalFormatting, 1, 2, 3, 'a');
        shouldThrow(conditionalFormatting.addRange, {}, 1, 2, 3, 4);

        expect(conditionalFormatting.addRange(1, 2, 3, 4)).toBe(conditionalFormatting);
    });

    it('addRule adds a rule', () => {
        shouldThrow(conditionalFormatting.addRule, conditionalFormatting, xl.CFORMAT_BEGINWITH, 'a');
        shouldThrow(conditionalFormatting.addRule, {}, xl.CFORMAT_BEGINWITH, conditionalFormat);

        expect(conditionalFormatting.addRule(xl.CFORMAT_BEGINWITH, conditionalFormat)).toBe(conditionalFormatting);
        expect(conditionalFormatting.addRule(xl.CFORMAT_BEGINWITH, conditionalFormat, 'a')).toBe(conditionalFormatting);
        expect(conditionalFormatting.addRule(xl.CFORMAT_BEGINWITH, conditionalFormat, 'a', false)).toBe(
            conditionalFormatting
        );
    });

    it('addTopRule adds a top rule', () => {
        shouldThrow(conditionalFormatting.addTopRule, conditionalFormatting, 'invalid', 10);
        shouldThrow(conditionalFormatting.addTopRule, {}, conditionalFormat, 10);

        expect(conditionalFormatting.addTopRule(conditionalFormat, 10)).toBe(conditionalFormatting);
        expect(conditionalFormatting.addTopRule(conditionalFormat, 10, true)).toBe(conditionalFormatting);
        expect(conditionalFormatting.addTopRule(conditionalFormat, 10, true, true)).toBe(conditionalFormatting);
        expect(conditionalFormatting.addTopRule(conditionalFormat, 10, true, true, true)).toBe(conditionalFormatting);
    });

    it('addOpNumRule adds a numeric operator rule', () => {
        shouldThrow(conditionalFormatting.addOpNumRule, conditionalFormatting, xl.CFOPERATOR_LESSTHAN, 'invalid', 10);
        shouldThrow(conditionalFormatting.addOpNumRule, {}, xl.CFOPERATOR_LESSTHAN, conditionalFormat, 10);

        expect(conditionalFormatting.addOpNumRule(xl.CFOPERATOR_LESSTHAN, conditionalFormat, 10)).toBe(
            conditionalFormatting
        );
        expect(conditionalFormatting.addOpNumRule(xl.CFOPERATOR_LESSTHAN, conditionalFormat, 10, 20)).toBe(
            conditionalFormatting
        );
        expect(conditionalFormatting.addOpNumRule(xl.CFOPERATOR_LESSTHAN, conditionalFormat, 10, 0, true)).toBe(
            conditionalFormatting
        );
    });

    it('addOpStrRule adds a string operator rule', () => {
        shouldThrow(
            conditionalFormatting.addOpStrRule,
            conditionalFormatting,
            xl.CFOPERATOR_LESSTHAN,
            'invalid',
            'test'
        );
        shouldThrow(conditionalFormatting.addOpStrRule, {}, xl.CFOPERATOR_LESSTHAN, conditionalFormat, 'test');

        expect(conditionalFormatting.addOpStrRule(xl.CFOPERATOR_LESSTHAN, conditionalFormat, 'test')).toBe(
            conditionalFormatting
        );
        expect(conditionalFormatting.addOpStrRule(xl.CFOPERATOR_LESSTHAN, conditionalFormat, 'a', 'z')).toBe(
            conditionalFormatting
        );
        expect(conditionalFormatting.addOpStrRule(xl.CFOPERATOR_LESSTHAN, conditionalFormat, 'test', '', true)).toBe(
            conditionalFormatting
        );
    });

    it('addAboveAverageRule adds an above average rule', () => {
        shouldThrow(conditionalFormatting.addAboveAverageRule, conditionalFormatting, 'invalid');
        shouldThrow(conditionalFormatting.addAboveAverageRule, {}, conditionalFormat);

        expect(conditionalFormatting.addAboveAverageRule(conditionalFormat)).toBe(conditionalFormatting);
        expect(conditionalFormatting.addAboveAverageRule(conditionalFormat, false)).toBe(conditionalFormatting);
        expect(conditionalFormatting.addAboveAverageRule(conditionalFormat, true, true)).toBe(conditionalFormatting);
        expect(conditionalFormatting.addAboveAverageRule(conditionalFormat, true, false, 1)).toBe(
            conditionalFormatting
        );
        expect(conditionalFormatting.addAboveAverageRule(conditionalFormat, true, false, 1, true)).toBe(
            conditionalFormatting
        );
    });

    it('addTimePeriodRule adds a time period rule', () => {
        shouldThrow(conditionalFormatting.addTimePeriodRule, conditionalFormatting, 'invalid', xl.CFTP_LAST7DAYS);
        shouldThrow(conditionalFormatting.addTimePeriodRule, {}, conditionalFormat, xl.CFTP_LAST7DAYS);

        expect(conditionalFormatting.addTimePeriodRule(conditionalFormat, xl.CFTP_LAST7DAYS)).toBe(
            conditionalFormatting
        );
        expect(conditionalFormatting.addTimePeriodRule(conditionalFormat, xl.CFTP_LAST7DAYS, true)).toBe(
            conditionalFormatting
        );
    });

    it('add2ColorScaleRule adds a 2-color scale rule', () => {
        shouldThrow(conditionalFormatting.add2ColorScaleRule, conditionalFormatting, 'invalid', xl.COLOR_RED);
        shouldThrow(conditionalFormatting.add2ColorScaleRule, {}, xl.COLOR_GREEN, xl.COLOR_RED);

        expect(conditionalFormatting.add2ColorScaleRule(xl.COLOR_GREEN, xl.COLOR_RED)).toBe(conditionalFormatting);
        expect(
            conditionalFormatting.add2ColorScaleRule(
                xl.COLOR_GREEN,
                xl.COLOR_RED,
                xl.CFVO_MIN,
                0,
                xl.CFVO_MAX,
                100,
                true
            )
        ).toBe(conditionalFormatting);
    });

    it('add2ColorScaleFormulaRule adds a 2-color scale formula rule', () => {
        shouldThrow(conditionalFormatting.add2ColorScaleFormulaRule, conditionalFormatting, 'invalid', xl.COLOR_RED);
        shouldThrow(conditionalFormatting.add2ColorScaleFormulaRule, {}, xl.COLOR_GREEN, xl.COLOR_RED);

        expect(conditionalFormatting.add2ColorScaleFormulaRule(xl.COLOR_GREEN, xl.COLOR_RED)).toBe(
            conditionalFormatting
        );
        expect(
            conditionalFormatting.add2ColorScaleFormulaRule(
                xl.COLOR_GREEN,
                xl.COLOR_RED,
                xl.CFVO_FORMULA,
                'A1',
                xl.CFVO_FORMULA,
                'B1',
                true
            )
        ).toBe(conditionalFormatting);
    });

    it('add3ColorScaleRule adds a 3-color scale rule', () => {
        shouldThrow(
            conditionalFormatting.add3ColorScaleRule,
            conditionalFormatting,
            'invalid',
            xl.COLOR_YELLOW,
            xl.COLOR_RED
        );
        shouldThrow(conditionalFormatting.add3ColorScaleRule, {}, xl.COLOR_GREEN, xl.COLOR_YELLOW, xl.COLOR_RED);

        expect(conditionalFormatting.add3ColorScaleRule(xl.COLOR_GREEN, xl.COLOR_YELLOW, xl.COLOR_RED)).toBe(
            conditionalFormatting
        );
        expect(
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
                true
            )
        ).toBe(conditionalFormatting);
    });

    it('add3ColorScaleFormulaRule adds a 3-color scale formula rule', () => {
        shouldThrow(
            conditionalFormatting.add3ColorScaleFormulaRule,
            conditionalFormatting,
            'invalid',
            xl.COLOR_YELLOW,
            xl.COLOR_RED
        );
        shouldThrow(conditionalFormatting.add3ColorScaleFormulaRule, {}, xl.COLOR_GREEN, xl.COLOR_YELLOW, xl.COLOR_RED);

        expect(conditionalFormatting.add3ColorScaleFormulaRule(xl.COLOR_GREEN, xl.COLOR_YELLOW, xl.COLOR_RED)).toBe(
            conditionalFormatting
        );
        expect(
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
                true
            )
        ).toBe(conditionalFormatting);
    });
});
