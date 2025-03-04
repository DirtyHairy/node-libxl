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

        expect(conditionalFormat.font()).toBeInstanceOf(xl.Font);
    });

    it('numFormat and setNumFormat manage the number format', () => {
        shouldThrow(conditionalFormat.setNumFormat, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setNumFormat, {}, xl.NUMFORMAT_GENERAL);
        shouldThrow(conditionalFormat.numFormat, {});

        expect(conditionalFormat.setNumFormat(xl.NUMFORMAT_GENERAL)).toBe(conditionalFormat);
        expect(conditionalFormat.numFormat()).toBe(xl.NUMFORMAT_GENERAL);
    });

    it('customNumFormat and setCustomNumFormat manage the number format', () => {
        shouldThrow(conditionalFormat.setCustomNumFormat, conditionalFormat, 10);
        shouldThrow(conditionalFormat.setCustomNumFormat, {}, 'abc');
        shouldThrow(conditionalFormat.customNumFormat, {});

        expect(conditionalFormat.setCustomNumFormat('abc')).toBe(conditionalFormat);
        expect(conditionalFormat.customNumFormat()).toBe('abc');
    });

    it('setBorder sets the border', () => {
        shouldThrow(conditionalFormat.setBorder, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setBorder, {}, xl.BORDERSTYLE_MEDIUM);
        shouldThrow(conditionalFormat.setBorder, {});

        expect(conditionalFormat.setBorder(xl.BORDERSTYLE_MEDIUM)).toBe(conditionalFormat);
        expect(conditionalFormat.setBorder()).toBe(conditionalFormat);
    });

    it('setBorderColor sets the border color', () => {
        shouldThrow(conditionalFormat.setBorderColor, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setBorderColor, {}, xl.COLOR_RED);

        expect(conditionalFormat.setBorderColor(xl.COLOR_RED)).toBe(conditionalFormat);
    });

    it('borderLeft and setBorderLeft manage left border', () => {
        shouldThrow(conditionalFormat.setBorderLeft, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setBorderLeft, {}, xl.BORDERSTYLE_MEDIUM);
        shouldThrow(conditionalFormat.borderLeft, {});

        expect(conditionalFormat.setBorderLeft()).toBe(conditionalFormat);
        expect(conditionalFormat.setBorderLeft(xl.BORDERSTYLE_MEDIUM)).toBe(conditionalFormat);
        expect(conditionalFormat.borderLeft()).toBe(xl.BORDERSTYLE_MEDIUM);
    });

    it('borderRight and setBorderRight manage right border', () => {
        shouldThrow(conditionalFormat.setBorderRight, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setBorderRight, {}, xl.BORDERSTYLE_MEDIUM);
        shouldThrow(conditionalFormat.borderRight, {});

        expect(conditionalFormat.setBorderRight()).toBe(conditionalFormat);
        expect(conditionalFormat.setBorderRight(xl.BORDERSTYLE_MEDIUM)).toBe(conditionalFormat);
        expect(conditionalFormat.borderRight()).toBe(xl.BORDERSTYLE_MEDIUM);
    });

    it('borderTop and setBorderTop manage top border', () => {
        shouldThrow(conditionalFormat.setBorderTop, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setBorderTop, {}, xl.BORDERSTYLE_MEDIUM);
        shouldThrow(conditionalFormat.borderTop, {});

        expect(conditionalFormat.setBorderTop()).toBe(conditionalFormat);
        expect(conditionalFormat.setBorderTop(xl.BORDERSTYLE_MEDIUM)).toBe(conditionalFormat);
        expect(conditionalFormat.borderTop()).toBe(xl.BORDERSTYLE_MEDIUM);
    });

    it('borderBottom and setBorderBottom manage bottom border', () => {
        shouldThrow(conditionalFormat.setBorderBottom, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setBorderBottom, {}, xl.BORDERSTYLE_MEDIUM);
        shouldThrow(conditionalFormat.borderBottom, {});

        expect(conditionalFormat.setBorderBottom()).toBe(conditionalFormat);
        expect(conditionalFormat.setBorderBottom(xl.BORDERSTYLE_MEDIUM)).toBe(conditionalFormat);
        expect(conditionalFormat.borderBottom()).toBe(xl.BORDERSTYLE_MEDIUM);
    });

    it('borderLeftColor and setBorderLeftColor manage left border color', () => {
        shouldThrow(conditionalFormat.borderLeftColor, {});
        shouldThrow(conditionalFormat.setBorderLeftColor, conditionalFormat, 'abc');

        expect(conditionalFormat.setBorderLeftColor(xl.COLOR_RED)).toBe(conditionalFormat);
        expect(conditionalFormat.borderLeftColor()).toBe(xl.COLOR_RED);
    });

    it('borderRightColor and setBorderRightColor manage right border color', () => {
        shouldThrow(conditionalFormat.borderRightColor, {});
        shouldThrow(conditionalFormat.setBorderRightColor, conditionalFormat, 'abc');

        expect(conditionalFormat.setBorderRightColor(xl.COLOR_GREEN)).toBe(conditionalFormat);
        expect(conditionalFormat.borderRightColor()).toBe(xl.COLOR_GREEN);
    });

    it('borderTopColor and setBorderTopColor manage top border color', () => {
        shouldThrow(conditionalFormat.borderTopColor, {});
        shouldThrow(conditionalFormat.setBorderTopColor, conditionalFormat, 'abc');

        expect(conditionalFormat.setBorderTopColor(xl.COLOR_BLUE)).toBe(conditionalFormat);
        expect(conditionalFormat.borderTopColor()).toBe(xl.COLOR_BLUE);
    });

    it('borderBottomColor and setBorderBottomColor manage bottom border color', () => {
        shouldThrow(conditionalFormat.borderBottomColor, {});
        shouldThrow(conditionalFormat.setBorderBottomColor, conditionalFormat, 'abc');

        expect(conditionalFormat.setBorderBottomColor(xl.COLOR_BLACK)).toBe(conditionalFormat);
        expect(conditionalFormat.borderBottomColor()).toBe(xl.COLOR_BLACK);
    });

    it('fillPattern and setFillPattern manage fill pattern', () => {
        shouldThrow(conditionalFormat.setFillPattern, conditionalFormat, 'abc');
        shouldThrow(conditionalFormat.setFillPattern, {}, xl.FILLPATTERN_SOLID);
        shouldThrow(conditionalFormat.fillPattern, {});

        expect(conditionalFormat.setFillPattern(xl.FILLPATTERN_SOLID)).toBe(conditionalFormat);
        expect(conditionalFormat.fillPattern()).toBe(xl.FILLPATTERN_SOLID);
    });

    it('patternForegroundColor and setPatternForegroundColor manage pattern foreground color', () => {
        shouldThrow(conditionalFormat.patternForegroundColor, {});
        shouldThrow(conditionalFormat.setPatternForegroundColor, conditionalFormat, 'abc');

        expect(conditionalFormat.setPatternForegroundColor(xl.COLOR_RED)).toBe(conditionalFormat);
        expect(conditionalFormat.patternForegroundColor()).toBe(xl.COLOR_RED);
    });

    it('patternBackgroundColor and setPatternBackgroundColor manage pattern background color', () => {
        shouldThrow(conditionalFormat.patternBackgroundColor, {});
        shouldThrow(conditionalFormat.setPatternBackgroundColor, conditionalFormat, 'abc');

        expect(conditionalFormat.setPatternBackgroundColor(xl.COLOR_BLUE)).toBe(conditionalFormat);
        expect(conditionalFormat.patternBackgroundColor()).toBe(xl.COLOR_BLUE);
    });
});
