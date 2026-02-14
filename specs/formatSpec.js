import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import util from 'util';
import xl from '../lib/libxl.js';

describe('The format class', () => {
    const book = new xl.Book(xl.BOOK_TYPE_XLS),
        format = book.addFormat(),
        font = book.addFont().setName('times');

    it('format.font returns the font', () => {
        assert.throws(() => format.font.call({}));
        assert.strictEqual(format.font() instanceof xl.Font, true);
    });

    it('format.setFont sets the font', () => {
        assert.throws(() => format.setFont.call(format, {}));
        assert.throws(() => format.setFont.call({}, font));
        assert.strictEqual(format.setFont(font), format);
        assert.strictEqual(format.font().name(), 'times');
    });

    it('format.numFormat returns the number format', () => {
        assert.throws(() => format.numFormat.call({}));
        format.numFormat();
    });

    it('format.setNumFormat sets the number format', () => {
        assert.throws(() => format.setNumFormat.call(format, 'a'));
        assert.throws(() => format.setNumFormat.call({}, xl.NUMFORMAT_DATE));
        assert.strictEqual(format.setNumFormat(xl.NUMFORMAT_DATE), format);
        assert.strictEqual(format.numFormat(), xl.NUMFORMAT_DATE);
        assert.strictEqual(format.setNumFormat(xl.NUMFORMAT_PERCENT).numFormat(), xl.NUMFORMAT_PERCENT);
    });

    it('format.alignH returns horizontal align', () => {
        assert.throws(() => format.alignH.call({}));
        format.alignH();
    });

    it('format.setAlignH sets horizontal aligh', () => {
        assert.throws(() => format.setAlignH.call(format, 'a'));
        assert.throws(() => format.setAlignH.call({}, xl.ALIGNH_LEFT));
        assert.strictEqual(format.setAlignH(xl.ALIGNH_LEFT), format);
        assert.strictEqual(format.alignH(), xl.ALIGNH_LEFT);
        assert.strictEqual(format.setAlignH(xl.ALIGNH_RIGHT).alignH(), xl.ALIGNH_RIGHT);
    });

    it('format.alignV returns vertical align', () => {
        assert.throws(() => format.alignV.call({}));
        format.alignV();
    });

    it('format.setAlignV sets vertical align', () => {
        assert.throws(() => format.setAlignV.call(format, 'a'));
        assert.throws(() => format.setAlignV.call({}, xl.ALIGNV_CENTER));
        assert.strictEqual(format.setAlignV(xl.ALIGNV_CENTER), format);
        assert.strictEqual(format.alignV(), xl.ALIGNV_CENTER);
        assert.strictEqual(format.setAlignV(xl.ALIGNV_BOTTOM).alignV(), xl.ALIGNV_BOTTOM);
    });

    it('format.wrap gets wrap', () => {
        assert.throws(() => format.wrap.call({}));
        format.wrap();
    });

    it('format.setWrap sets wrap', () => {
        assert.throws(() => format.setWrap.call(format, 1));
        assert.throws(() => format.setWrap.call({}, true));
        assert.strictEqual(format.setWrap(true), format);
        assert.strictEqual(format.wrap(), true);
        assert.strictEqual(format.setWrap(false).wrap(), false);
    });

    it('format.rotation gets text rotation', () => {
        assert.throws(() => format.rotation.call({}));
        format.rotation();
    });

    it('format.setRotation sets text rotation', () => {
        assert.throws(() => format.setRotation.call(format, 'a'));
        assert.throws(() => format.setRotation.call({}, 34));
        assert.strictEqual(format.setRotation(34), format);
        assert.strictEqual(format.rotation(), 34);
        assert.strictEqual(format.setRotation(78).rotation(), 78);
    });

    it('format.indent gets text indentation', () => {
        assert.throws(() => format.indent.call({}));
        format.indent();
    });

    it('format.setIndent sets text indentation', () => {
        assert.throws(() => format.setIndent.call(format, 'a'));
        assert.throws(() => format.setIndent.call({}, 9));
        assert.strictEqual(format.setIndent(9), format);
        assert.strictEqual(format.indent(), 9);
        assert.strictEqual(format.setIndent(11).indent(), 11);
    });

    it('format.shrinkToFit checks shrink-to-fit', () => {
        assert.throws(() => format.shrinkToFit.call({}));
        format.shrinkToFit();
    });

    it('format.setShrinkToFit toggles shrink-to-fit', () => {
        assert.throws(() => format.setShrinkToFit.call(format, 1));
        assert.throws(() => format.setShrinkToFit.call({}, true));
        assert.strictEqual(format.setShrinkToFit(true), format);
        assert.strictEqual(format.shrinkToFit(), true);
        assert.strictEqual(format.setShrinkToFit(false).shrinkToFit(), false);
    });

    it('format.setBorder sets all borders', () => {
        const expectAllBordersToBe = (type) => {
            ['Left', 'Right', 'Bottom', 'Top'].forEach((border) => {
                assert.strictEqual(format['border' + border](), type);
            });
        };

        assert.throws(() => format.setBorder.format.call(undefined, 'a'));
        assert.throws(() => format.setBorder.call({}, xl.BORDERSTYLE_THIN));
        assert.strictEqual(format.setBorder(xl.BORDERSTYLE_THIN), format);
        expectAllBordersToBe(xl.BORDERSTYLE_THIN);
        assert.strictEqual(format.setBorder(xl.BORDERSTYLE_MEDIUM), format);
        expectAllBordersToBe(xl.BORDERSTYLE_MEDIUM);
    });

    it('format.setBorderColor sets the color of all borders', () => {
        const expectAllBordersToBe = (color) => {
            ['Left', 'Right', 'Bottom', 'Top'].forEach((border) => {
                assert.strictEqual(format['border' + border + 'Color'](), color);
            });
        };

        assert.throws(() => format.setBorderColor.call(format, 'a'));
        assert.throws(() => format.setBorderColor.call({}, xl.COLOR_YELLOW));
        assert.strictEqual(format.setBorderColor(xl.COLOR_YELLOW), format);
        expectAllBordersToBe(xl.COLOR_YELLOW);
        assert.strictEqual(format.setBorderColor(xl.COLOR_GREEN), format);
        expectAllBordersToBe(xl.COLOR_GREEN);
    });

    ['Left', 'Right', 'Top', 'Bottom'].forEach((border) => {
        const styleGetter = 'border' + border,
            styleSetter = 'setBorder' + border,
            colorGetter = 'border' + border + 'Color',
            colorSetter = 'setBorder' + border + 'Color';

        it(util.format('format.%s returns the %s border style', styleGetter, border.toLowerCase()), () => {
            assert.throws(() => format[styleGetter].call({}));
            format[styleGetter]();
        });

        it(util.format('format.%s sets the %s border style', styleSetter, border.toLowerCase()), () => {
            assert.throws(() => format[styleSetter].call(format, 'a'));
            assert.throws(() => format[styleSetter].call({}, xl.BORDERSTYLE_THIN));
            assert.strictEqual(format[styleSetter](xl.BORDERSTYLE_THIN), format);
            assert.strictEqual(format[styleGetter](), xl.BORDERSTYLE_THIN);
            assert.strictEqual(format[styleSetter](xl.BORDERSTYLE_MEDIUM)[styleGetter](), xl.BORDERSTYLE_MEDIUM);
        });
    });

    ['Left', 'Right', 'Top', 'Bottom', 'Diagonal'].forEach((border) => {
        const colorGetter = 'border' + border + 'Color',
            colorSetter = 'setBorder' + border + 'Color';

        it(util.format('format.%s returns the %s border color', colorGetter, border.toLowerCase()), () => {
            assert.throws(() => format[colorGetter]().call({}));
            format[colorGetter]();
        });

        it(util.format('format.%s sets the %s border color', colorSetter, border.toLowerCase()), () => {
            assert.throws(() => format[colorSetter].call(format, 'a'));
            assert.throws(() => format[colorSetter].call({}, xl.COLOR_YELLOW));
            assert.strictEqual(format[colorSetter](xl.COLOR_YELLOW), format);
            assert.strictEqual(format[colorGetter](), xl.COLOR_YELLOW);
            assert.strictEqual(format[colorSetter](xl.COLOR_GREEN)[colorGetter](), xl.COLOR_GREEN);
        });
    });

    it('format.borderDiagonal returns diagonal border style', () => {
        assert.throws(() => format.borderDiagonal.call({}));
        format.borderDiagonal();
    });

    it('format.setBorderDiagonal sets diagonal border style', () => {
        assert.throws(() => format.setBorderDiagonal.call(format, 'a'));
        assert.throws(() => format.setBorderDiagonal.call({}, xl.BORDERDIAGONAL_DOWN));
        assert.strictEqual(format.setBorderDiagonal(xl.BORDERDIAGONAL_DOWN), format);
        assert.strictEqual(format.borderDiagonal(), xl.BORDERDIAGONAL_DOWN);
        assert.strictEqual(format.setBorderDiagonal(xl.BORDERDIAGONAL_UP).borderDiagonal(), xl.BORDERDIAGONAL_UP);
    });

    it('format.fillPattern returns the fill pattern', () => {
        assert.throws(() => format.fillPattern().call({}));
        format.fillPattern();
    });

    it('format.setFillPattern sets the fill pattern', () => {
        assert.throws(() => format.setFillPattern.call(format, 'a'));
        assert.throws(() => format.setFillPattern.call({}, xl.FILLPATTERN_GRAY50));
        assert.strictEqual(format.setFillPattern(xl.FILLPATTERN_GRAY50), format);
        assert.strictEqual(format.fillPattern(), xl.FILLPATTERN_GRAY50);
        assert.strictEqual(format.setFillPattern(xl.FILLPATTERN_GRAY75).fillPattern(), xl.FILLPATTERN_GRAY75);
    });

    ['Foreground', 'Background'].forEach((layer) => {
        const setter = 'setPattern' + layer + 'Color',
            getter = 'pattern' + layer + 'Color';

        if (
            (util.format('format.%s returns %s pattern color', getter, layer.toLowerCase()),
            () => {
                assert.throws(() => format[getter].call({}));
                format[getter]();
            })
        );

        it(util.format('format.%s sets %s pattern color', setter, layer.toLowerCase()), () => {
            assert.throws(() => format[setter].call(format, 'a'));
            assert.throws(() => format[setter].call({}, xl.COLOR_YELLOW));
            assert.strictEqual(format[setter](xl.COLOR_YELLOW), format);
            assert.strictEqual(format[getter](), xl.COLOR_YELLOW);
            assert.strictEqual(format[setter](xl.COLOR_GREEN)[getter](), xl.COLOR_GREEN);
        });
    });

    it('format.locked checks whether the format is locked', () => {
        assert.throws(() => format.locked.call({}));
        format.locked();
    });

    it('format.setLocked toggles a format locked', () => {
        assert.throws(() => format.setLocked.call(format, 1));
        assert.throws(() => format.setLocked.call({}, true));
        assert.strictEqual(format.setLocked(true), format);
        assert.strictEqual(format.locked(), true);
        assert.strictEqual(format.setLocked(false).locked(), false);
    });

    it('format.hidden checks whether the format is hidden', () => {
        assert.throws(() => format.hidden.call({}));
        format.hidden();
    });

    it('format.setHidden toggles a format hidden', () => {
        assert.throws(() => format.setHidden.call(format, 1));
        assert.throws(() => format.setHidden.call({}, true));
        assert.strictEqual(format.setHidden(true), format);
        assert.strictEqual(format.hidden(), true);
        assert.strictEqual(format.setHidden(false).hidden(), false);
    });
});
