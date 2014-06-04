var xl = require('../lib/libxl'),
    util = require('util'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('The format class', function() {
    var book = new xl.Book(xl.BOOK_TYPE_XLS),
        format = book.addFormat(),
        font = book.addFont().setName('times');

    it('format.font returns the font', function() {
        shouldThrow(format.font, {});
        expect(format.font() instanceof font.constructor).toBe(true);
    });

    it('format.setFont sets the font', function() {
        shouldThrow(format.setFont, format, {});
        shouldThrow(format.setFont, {}, font);
        expect(format.setFont(font)).toBe(format);
        expect(format.font().name()).toBe('times');
    });

    it('format.numFormat returns the number format', function() {
        shouldThrow(format.numFormat, {});
        format.numFormat();
    });

    it('format.setNumFormat sets the number format', function() {
        shouldThrow(format.setNumFormat, format, 'a');
        shouldThrow(format.setNumFormat, {}, xl.NUMFORMAT_DATE);
        expect(format.setNumFormat(xl.NUMFORMAT_DATE)).toBe(format);
        expect(format.numFormat()).toBe(xl.NUMFORMAT_DATE);
        expect(format.setNumFormat(xl.NUMFORMAT_PERCENT).numFormat()).toBe(
            xl.NUMFORMAT_PERCENT);
    });

    it('format.alignH returns horizontal align', function() {
        shouldThrow(format.alignH, {});
        format.alignH();
    });

    it('format.setAlignH sets horizontal aligh', function() {
        shouldThrow(format.setAlignH, format, 'a');
        shouldThrow(format.setAlignH, {}, xl.ALIGNH_LEFT);
        expect(format.setAlignH(xl.ALIGNH_LEFT)).toBe(format);
        expect(format.alignH()).toBe(xl.ALIGNH_LEFT);
        expect(format.setAlignH(xl.ALIGNH_RIGHT).alignH()).toBe(xl.ALIGNH_RIGHT);
    });

    it('format.alignV returns vertical align', function() {
        shouldThrow(format.alignV, {});
        format.alignV();
    });

    it('format.setAlignV sets vertical align', function() {
        shouldThrow(format.setAlignV, format, 'a');
        shouldThrow(format.setAlignV, {}, xl.ALIGNV_CENTER);
        expect(format.setAlignV(xl.ALIGNV_CENTER)).toBe(format);
        expect(format.alignV()).toBe(xl.ALIGNV_CENTER);
        expect(format.setAlignV(xl.ALIGNV_BOTTOM).alignV()).toBe(xl.ALIGNV_BOTTOM);
    });

    it('format.wrap gets wrap', function() {
        shouldThrow(format.wrap, {});
        format.wrap();
    });

    it('format.setWrap sets wrap', function() {
        shouldThrow(format.setWrap, format, 1);
        shouldThrow(format.setWrap, {}, true);
        expect(format.setWrap(true)).toBe(format);
        expect(format.wrap()).toBe(true);
        expect(format.setWrap(false).wrap()).toBe(false);
    });

    it('format.rotation gets text rotation', function() {
        shouldThrow(format.rotation, {});
        format.rotation();
    });

    it('format.setRotation sets text rotation', function() {
        shouldThrow(format.setRotation, format, 'a');
        shouldThrow(format.setRotation, {}, 34);
        expect(format.setRotation(34)).toBe(format);
        expect(format.rotation()).toBe(34);
        expect(format.setRotation(78).rotation()).toBe(78);
    });

    it('format.indent gets text indentation', function() {
        shouldThrow(format.indent, {});
        format.indent();
    });

    it('format.setIndent sets text indentation', function() {
        shouldThrow(format.setIndent, format, 'a');
        shouldThrow(format.setIndent, {}, 9);
        expect(format.setIndent(9)).toBe(format);
        expect(format.indent()).toBe(9);
        expect(format.setIndent(11).indent()).toBe(11);
    });

    it('format.shrinkToFit checks shrink-to-fit', function() {
        shouldThrow(format.shrinkToFit, {});
        format.shrinkToFit();
    });

    it('format.setShrinkToFit toggles shrink-to-fit', function() {
        shouldThrow(format.setShrinkToFit, format, 1);
        shouldThrow(format.setShrinkToFit, {}, true);
        expect(format.setShrinkToFit(true)).toBe(format);
        expect(format.shrinkToFit()).toBe(true);
        expect(format.setShrinkToFit(false).shrinkToFit()).toBe(false);
    });

    it('format.setBorder sets all borders', function() {
        function expectAllBordersToBe(type) {
            ['Left', 'Right', 'Bottom', 'Top'].forEach(function(border) {
                expect(format['border' + border]()).toBe(type);
            });
        }

        shouldThrow(format.setBorder. format, 'a');
        shouldThrow(format.setBorder, {}, xl.BORDERSTYLE_THIN);
        expect(format.setBorder(xl.BORDERSTYLE_THIN)).toBe(format);
        expectAllBordersToBe(xl.BORDERSTYLE_THIN);
        expect(format.setBorder(xl.BORDERSTYLE_MEDIUM)).toBe(format);
        expectAllBordersToBe(xl.BORDERSTYLE_MEDIUM);
    });

    it('format.setBorderColor sets the color of all borders', function() {
        function expectAllBordersToBe(color) {
            ['Left', 'Right', 'Bottom', 'Top'].forEach(function(border) {
                expect(format['border' + border + 'Color']()).toBe(color);
            });
        }

        shouldThrow(format.setBorderColor, format, 'a');
        shouldThrow(format.setBorderColor, {}, xl.COLOR_YELLOW);
        expect(format.setBorderColor(xl.COLOR_YELLOW)).toBe(format);
        expectAllBordersToBe(xl.COLOR_YELLOW);
        expect(format.setBorderColor(xl.COLOR_GREEN)).toBe(format);
        expectAllBordersToBe(xl.COLOR_GREEN);
    });

    ['Left', 'Right', 'Top', 'Bottom'].forEach(function(border) {
        var styleGetter = 'border' + border,
            styleSetter = 'setBorder' + border,
            colorGetter = 'border' + border + 'Color',
            colorSetter = 'setBorder' + border + 'Color';
        
        it(util.format('format.%s returns the %s border style',
            styleGetter, border.toLowerCase()),  function()
        {
            shouldThrow(format[styleGetter], {});
            format[styleGetter]();
        });

        it(util.format('format.%s sets the %s border style',
            styleSetter, border.toLowerCase()), function()
        {
            shouldThrow(format[styleSetter], format, 'a');
            shouldThrow(format[styleSetter], {}, xl.BORDERSTYLE_THIN);
            expect(format[styleSetter](xl.BORDERSTYLE_THIN)).toBe(format);
            expect(format[styleGetter]()).toBe(xl.BORDERSTYLE_THIN);
            expect(format[styleSetter](xl.BORDERSTYLE_MEDIUM)[styleGetter]())
                .toBe(xl.BORDERSTYLE_MEDIUM);
        });
    });


    ['Left', 'Right', 'Top', 'Bottom', 'Diagonal'].forEach(function(border) {
        var colorGetter = 'border' + border + 'Color',
            colorSetter = 'setBorder' + border + 'Color';
 
        it(util.format('format.%s returns the %s border color',
            colorGetter, border.toLowerCase()), function()
        {
            shouldThrow(format[colorGetter](), {});
            format[colorGetter]();
        });

        it(util.format('format.%s sets the %s border color',
            colorSetter, border.toLowerCase()), function()
        {
            shouldThrow(format[colorSetter], format, 'a');
            shouldThrow(format[colorSetter], {}, xl.COLOR_YELLOW);
            expect(format[colorSetter](xl.COLOR_YELLOW)).toBe(format);
            expect(format[colorGetter]()).toBe(xl.COLOR_YELLOW);
            expect(format[colorSetter](xl.COLOR_GREEN)[colorGetter]())
                .toBe(xl.COLOR_GREEN);
        });
    });

    it('format.borderDiagonal returns diagonal border style', function() {
        shouldThrow(format.borderDiagonal, {});
        format.borderDiagonal();
    });

    it('format.setBorderDiagonal sets diagonal border style', function() {
        shouldThrow(format.setBorderDiagonal, format, 'a');
        shouldThrow(format.setBorderDiagonal, {}, xl.BORDERDIAGONAL_DOWN);
        expect(format.setBorderDiagonal(xl.BORDERDIAGONAL_DOWN)).toBe(format);
        expect(format.borderDiagonal()).toBe(xl.BORDERDIAGONAL_DOWN);
        expect(format.setBorderDiagonal(xl.BORDERDIAGONAL_UP).borderDiagonal())
            .toBe(xl.BORDERDIAGONAL_UP);
    });

    it('format.fillPattern returns the fill pattern', function() {
        shouldThrow(format.fillPattern(), {});
        format.fillPattern();
    });

    it('format.setFillPattern sets the fill pattern', function() {
        shouldThrow(format.setFillPattern, format, 'a');
        shouldThrow(format.setFillPattern, {}, xl.FILLPATTERN_GRAY50);
        expect(format.setFillPattern(xl.FILLPATTERN_GRAY50)).toBe(format);
        expect(format.fillPattern()).toBe(xl.FILLPATTERN_GRAY50);
        expect(format.setFillPattern(xl.FILLPATTERN_GRAY75).fillPattern())
            .toBe(xl.FILLPATTERN_GRAY75);
    });

    ['Foreground', 'Background'].forEach(function(layer) {
        var setter = 'setPattern' + layer + 'Color',
            getter = 'pattern' + layer + 'Color';

        if(util.format('format.%s returns %s pattern color',
            getter, layer.toLowerCase()), function()
        {
            shouldThrow(format[getter], {});
            format[getter]();
        });

        it(util.format('format.%s sets %s pattern color',
            setter, layer.toLowerCase()), function()
        {
            shouldThrow(format[setter], format, 'a');
            shouldThrow(format[setter], {}, xl.COLOR_YELLOW);
            expect(format[setter](xl.COLOR_YELLOW)).toBe(format);
            expect(format[getter]()).toBe(xl.COLOR_YELLOW);
            expect(format[setter](xl.COLOR_GREEN)[getter]())
                .toBe(xl.COLOR_GREEN);
        });
    });

    it('format.locked checks whether the format is locked', function() {
        shouldThrow(format.locked, {});
        format.locked();
    });

    it('format.setLocked toggles a format locked', function() {
        shouldThrow(format.setLocked, format, 1);
        shouldThrow(format.setLocked, {}, true);
        expect(format.setLocked(true)).toBe(format);
        expect(format.locked()).toBe(true);
        expect(format.setLocked(false).locked()).toBe(false);
    });

    it('format.hidden checks whether the format is hidden', function() {
        shouldThrow(format.hidden, {});
        format.hidden();
    });

    it('format.setHidden toggles a format hidden', function() {
        shouldThrow(format.setHidden, format, 1);
        shouldThrow(format.setHidden, {}, true);
        expect(format.setHidden(true)).toBe(format);
        expect(format.hidden()).toBe(true);
        expect(format.setHidden(false).hidden()).toBe(false);
    });
});

