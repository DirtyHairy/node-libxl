const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
var xl = require('../lib/libxl'),
    util = require('util'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('The format class', function () {
    var book = new xl.Book(xl.BOOK_TYPE_XLS),
        format = book.addFormat(),
        font = book.addFont().setName('times');

    it('format.font returns the font', function () {
        shouldThrow(format.font, {});
        assert.strictEqual(format.font() instanceof xl.Font, true);
    });

    it('format.setFont sets the font', function () {
        shouldThrow(format.setFont, format, {});
        shouldThrow(format.setFont, {}, font);
        assert.strictEqual(format.setFont(font), format);
        assert.strictEqual(format.font().name(), 'times');
    });

    it('format.numFormat returns the number format', function () {
        shouldThrow(format.numFormat, {});
        format.numFormat();
    });

    it('format.setNumFormat sets the number format', function () {
        shouldThrow(format.setNumFormat, format, 'a');
        shouldThrow(format.setNumFormat, {}, xl.NUMFORMAT_DATE);
        assert.strictEqual(format.setNumFormat(xl.NUMFORMAT_DATE), format);
        assert.strictEqual(format.numFormat(), xl.NUMFORMAT_DATE);
        assert.strictEqual(format.setNumFormat(xl.NUMFORMAT_PERCENT).numFormat(), xl.NUMFORMAT_PERCENT);
    });

    it('format.alignH returns horizontal align', function () {
        shouldThrow(format.alignH, {});
        format.alignH();
    });

    it('format.setAlignH sets horizontal aligh', function () {
        shouldThrow(format.setAlignH, format, 'a');
        shouldThrow(format.setAlignH, {}, xl.ALIGNH_LEFT);
        assert.strictEqual(format.setAlignH(xl.ALIGNH_LEFT), format);
        assert.strictEqual(format.alignH(), xl.ALIGNH_LEFT);
        assert.strictEqual(format.setAlignH(xl.ALIGNH_RIGHT).alignH(), xl.ALIGNH_RIGHT);
    });

    it('format.alignV returns vertical align', function () {
        shouldThrow(format.alignV, {});
        format.alignV();
    });

    it('format.setAlignV sets vertical align', function () {
        shouldThrow(format.setAlignV, format, 'a');
        shouldThrow(format.setAlignV, {}, xl.ALIGNV_CENTER);
        assert.strictEqual(format.setAlignV(xl.ALIGNV_CENTER), format);
        assert.strictEqual(format.alignV(), xl.ALIGNV_CENTER);
        assert.strictEqual(format.setAlignV(xl.ALIGNV_BOTTOM).alignV(), xl.ALIGNV_BOTTOM);
    });

    it('format.wrap gets wrap', function () {
        shouldThrow(format.wrap, {});
        format.wrap();
    });

    it('format.setWrap sets wrap', function () {
        shouldThrow(format.setWrap, format, 1);
        shouldThrow(format.setWrap, {}, true);
        assert.strictEqual(format.setWrap(true), format);
        assert.strictEqual(format.wrap(), true);
        assert.strictEqual(format.setWrap(false).wrap(), false);
    });

    it('format.rotation gets text rotation', function () {
        shouldThrow(format.rotation, {});
        format.rotation();
    });

    it('format.setRotation sets text rotation', function () {
        shouldThrow(format.setRotation, format, 'a');
        shouldThrow(format.setRotation, {}, 34);
        assert.strictEqual(format.setRotation(34), format);
        assert.strictEqual(format.rotation(), 34);
        assert.strictEqual(format.setRotation(78).rotation(), 78);
    });

    it('format.indent gets text indentation', function () {
        shouldThrow(format.indent, {});
        format.indent();
    });

    it('format.setIndent sets text indentation', function () {
        shouldThrow(format.setIndent, format, 'a');
        shouldThrow(format.setIndent, {}, 9);
        assert.strictEqual(format.setIndent(9), format);
        assert.strictEqual(format.indent(), 9);
        assert.strictEqual(format.setIndent(11).indent(), 11);
    });

    it('format.shrinkToFit checks shrink-to-fit', function () {
        shouldThrow(format.shrinkToFit, {});
        format.shrinkToFit();
    });

    it('format.setShrinkToFit toggles shrink-to-fit', function () {
        shouldThrow(format.setShrinkToFit, format, 1);
        shouldThrow(format.setShrinkToFit, {}, true);
        assert.strictEqual(format.setShrinkToFit(true), format);
        assert.strictEqual(format.shrinkToFit(), true);
        assert.strictEqual(format.setShrinkToFit(false).shrinkToFit(), false);
    });

    it('format.setBorder sets all borders', function () {
        function expectAllBordersToBe(type) {
            ['Left', 'Right', 'Bottom', 'Top'].forEach(function (border) {
                assert.strictEqual(format['border' + border](), type);
            });
        }

        shouldThrow(format.setBorder.format, 'a');
        shouldThrow(format.setBorder, {}, xl.BORDERSTYLE_THIN);
        assert.strictEqual(format.setBorder(xl.BORDERSTYLE_THIN), format);
        expectAllBordersToBe(xl.BORDERSTYLE_THIN);
        assert.strictEqual(format.setBorder(xl.BORDERSTYLE_MEDIUM), format);
        expectAllBordersToBe(xl.BORDERSTYLE_MEDIUM);
    });

    it('format.setBorderColor sets the color of all borders', function () {
        function expectAllBordersToBe(color) {
            ['Left', 'Right', 'Bottom', 'Top'].forEach(function (border) {
                assert.strictEqual(format['border' + border + 'Color'](), color);
            });
        }

        shouldThrow(format.setBorderColor, format, 'a');
        shouldThrow(format.setBorderColor, {}, xl.COLOR_YELLOW);
        assert.strictEqual(format.setBorderColor(xl.COLOR_YELLOW), format);
        expectAllBordersToBe(xl.COLOR_YELLOW);
        assert.strictEqual(format.setBorderColor(xl.COLOR_GREEN), format);
        expectAllBordersToBe(xl.COLOR_GREEN);
    });

    ['Left', 'Right', 'Top', 'Bottom'].forEach(function (border) {
        var styleGetter = 'border' + border,
            styleSetter = 'setBorder' + border,
            colorGetter = 'border' + border + 'Color',
            colorSetter = 'setBorder' + border + 'Color';

        it(util.format('format.%s returns the %s border style', styleGetter, border.toLowerCase()), function () {
            shouldThrow(format[styleGetter], {});
            format[styleGetter]();
        });

        it(util.format('format.%s sets the %s border style', styleSetter, border.toLowerCase()), function () {
            shouldThrow(format[styleSetter], format, 'a');
            shouldThrow(format[styleSetter], {}, xl.BORDERSTYLE_THIN);
            assert.strictEqual(format[styleSetter](xl.BORDERSTYLE_THIN), format);
            assert.strictEqual(format[styleGetter](), xl.BORDERSTYLE_THIN);
            assert.strictEqual(format[styleSetter](xl.BORDERSTYLE_MEDIUM)[styleGetter](), xl.BORDERSTYLE_MEDIUM);
        });
    });

    ['Left', 'Right', 'Top', 'Bottom', 'Diagonal'].forEach(function (border) {
        var colorGetter = 'border' + border + 'Color',
            colorSetter = 'setBorder' + border + 'Color';

        it(util.format('format.%s returns the %s border color', colorGetter, border.toLowerCase()), function () {
            shouldThrow(format[colorGetter](), {});
            format[colorGetter]();
        });

        it(util.format('format.%s sets the %s border color', colorSetter, border.toLowerCase()), function () {
            shouldThrow(format[colorSetter], format, 'a');
            shouldThrow(format[colorSetter], {}, xl.COLOR_YELLOW);
            assert.strictEqual(format[colorSetter](xl.COLOR_YELLOW), format);
            assert.strictEqual(format[colorGetter](), xl.COLOR_YELLOW);
            assert.strictEqual(format[colorSetter](xl.COLOR_GREEN)[colorGetter](), xl.COLOR_GREEN);
        });
    });

    it('format.borderDiagonal returns diagonal border style', function () {
        shouldThrow(format.borderDiagonal, {});
        format.borderDiagonal();
    });

    it('format.setBorderDiagonal sets diagonal border style', function () {
        shouldThrow(format.setBorderDiagonal, format, 'a');
        shouldThrow(format.setBorderDiagonal, {}, xl.BORDERDIAGONAL_DOWN);
        assert.strictEqual(format.setBorderDiagonal(xl.BORDERDIAGONAL_DOWN), format);
        assert.strictEqual(format.borderDiagonal(), xl.BORDERDIAGONAL_DOWN);
        assert.strictEqual(format.setBorderDiagonal(xl.BORDERDIAGONAL_UP).borderDiagonal(), xl.BORDERDIAGONAL_UP);
    });

    it('format.fillPattern returns the fill pattern', function () {
        shouldThrow(format.fillPattern(), {});
        format.fillPattern();
    });

    it('format.setFillPattern sets the fill pattern', function () {
        shouldThrow(format.setFillPattern, format, 'a');
        shouldThrow(format.setFillPattern, {}, xl.FILLPATTERN_GRAY50);
        assert.strictEqual(format.setFillPattern(xl.FILLPATTERN_GRAY50), format);
        assert.strictEqual(format.fillPattern(), xl.FILLPATTERN_GRAY50);
        assert.strictEqual(format.setFillPattern(xl.FILLPATTERN_GRAY75).fillPattern(), xl.FILLPATTERN_GRAY75);
    });

    ['Foreground', 'Background'].forEach(function (layer) {
        var setter = 'setPattern' + layer + 'Color',
            getter = 'pattern' + layer + 'Color';

        if (
            (util.format('format.%s returns %s pattern color', getter, layer.toLowerCase()),
            function () {
                shouldThrow(format[getter], {});
                format[getter]();
            })
        );

        it(util.format('format.%s sets %s pattern color', setter, layer.toLowerCase()), function () {
            shouldThrow(format[setter], format, 'a');
            shouldThrow(format[setter], {}, xl.COLOR_YELLOW);
            assert.strictEqual(format[setter](xl.COLOR_YELLOW), format);
            assert.strictEqual(format[getter](), xl.COLOR_YELLOW);
            assert.strictEqual(format[setter](xl.COLOR_GREEN)[getter](), xl.COLOR_GREEN);
        });
    });

    it('format.locked checks whether the format is locked', function () {
        shouldThrow(format.locked, {});
        format.locked();
    });

    it('format.setLocked toggles a format locked', function () {
        shouldThrow(format.setLocked, format, 1);
        shouldThrow(format.setLocked, {}, true);
        assert.strictEqual(format.setLocked(true), format);
        assert.strictEqual(format.locked(), true);
        assert.strictEqual(format.setLocked(false).locked(), false);
    });

    it('format.hidden checks whether the format is hidden', function () {
        shouldThrow(format.hidden, {});
        format.hidden();
    });

    it('format.setHidden toggles a format hidden', function () {
        shouldThrow(format.setHidden, format, 1);
        shouldThrow(format.setHidden, {}, true);
        assert.strictEqual(format.setHidden(true), format);
        assert.strictEqual(format.hidden(), true);
        assert.strictEqual(format.setHidden(false).hidden(), false);
    });
});
