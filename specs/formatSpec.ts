import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import util from 'util';
import * as xl from '../lib/libxl';

describe('The format class', () => {
    const book = new xl.Book(xl.BOOK_TYPE_XLS),
        format = book.addFormat(),
        font = book.addFont().setName('times');

    it('format.font returns the font', () => {
        assert.throws(() => (format.font as any).call({}));
        assert.strictEqual(format.font() instanceof xl.Font, true);
    });

    it('format.setFont sets the font', () => {
        assert.throws(() => (format.setFont as any).call(format, {}));
        assert.throws(() => (format.setFont as any).call({}, font));
        assert.strictEqual(format.setFont(font), format);
        assert.strictEqual(format.font().name(), 'times');
    });

    it('format.numFormat returns the number format', () => {
        assert.throws(() => (format.numFormat as any).call({}));
        format.numFormat();
    });

    it('format.setNumFormat sets the number format', () => {
        assert.throws(() => (format.setNumFormat as any).call(format, 'a'));
        assert.throws(() => (format.setNumFormat as any).call({}, xl.NUMFORMAT_DATE));
        assert.strictEqual(format.setNumFormat(xl.NUMFORMAT_DATE), format);
        assert.strictEqual(format.numFormat(), xl.NUMFORMAT_DATE);
        assert.strictEqual(format.setNumFormat(xl.NUMFORMAT_PERCENT).numFormat(), xl.NUMFORMAT_PERCENT);
    });

    it('format.alignH returns horizontal align', () => {
        assert.throws(() => (format.alignH as any).call({}));
        format.alignH();
    });

    it('format.setAlignH sets horizontal aligh', () => {
        assert.throws(() => (format.setAlignH as any).call(format, 'a'));
        assert.throws(() => (format.setAlignH as any).call({}, xl.ALIGNH_LEFT));
        assert.strictEqual(format.setAlignH(xl.ALIGNH_LEFT), format);
        assert.strictEqual(format.alignH(), xl.ALIGNH_LEFT);
        assert.strictEqual(format.setAlignH(xl.ALIGNH_RIGHT).alignH(), xl.ALIGNH_RIGHT);
    });

    it('format.alignV returns vertical align', () => {
        assert.throws(() => (format.alignV as any).call({}));
        format.alignV();
    });

    it('format.setAlignV sets vertical align', () => {
        assert.throws(() => (format.setAlignV as any).call(format, 'a'));
        assert.throws(() => (format.setAlignV as any).call({}, xl.ALIGNV_CENTER));
        assert.strictEqual(format.setAlignV(xl.ALIGNV_CENTER), format);
        assert.strictEqual(format.alignV(), xl.ALIGNV_CENTER);
        assert.strictEqual(format.setAlignV(xl.ALIGNV_BOTTOM).alignV(), xl.ALIGNV_BOTTOM);
    });

    it('format.wrap gets wrap', () => {
        assert.throws(() => (format.wrap as any).call({}));
        format.wrap();
    });

    it('format.setWrap sets wrap', () => {
        assert.throws(() => (format.setWrap as any).call(format, 1));
        assert.throws(() => (format.setWrap as any).call({}, true));
        assert.strictEqual(format.setWrap(true), format);
        assert.strictEqual(format.wrap(), true);
        assert.strictEqual(format.setWrap(false).wrap(), false);
        assert.strictEqual(format.setWrap().wrap(), true);
    });

    it('format.rotation gets text rotation', () => {
        assert.throws(() => (format.rotation as any).call({}));
        format.rotation();
    });

    it('format.setRotation sets text rotation', () => {
        assert.throws(() => (format.setRotation as any).call(format, 'a'));
        assert.throws(() => (format.setRotation as any).call({}, 34));
        assert.strictEqual(format.setRotation(34), format);
        assert.strictEqual(format.rotation(), 34);
        assert.strictEqual(format.setRotation(78).rotation(), 78);
    });

    it('format.indent gets text indentation', () => {
        assert.throws(() => (format.indent as any).call({}));
        format.indent();
    });

    it('format.setIndent sets text indentation', () => {
        assert.throws(() => (format.setIndent as any).call(format, 'a'));
        assert.throws(() => (format.setIndent as any).call({}, 9));
        assert.strictEqual(format.setIndent(9), format);
        assert.strictEqual(format.indent(), 9);
        assert.strictEqual(format.setIndent(11).indent(), 11);
    });

    it('format.shrinkToFit checks shrink-to-fit', () => {
        assert.throws(() => (format.shrinkToFit as any).call({}));
        format.shrinkToFit();
    });

    it('format.setShrinkToFit toggles shrink-to-fit', () => {
        assert.throws(() => (format.setShrinkToFit as any).call(format, 1));
        assert.throws(() => (format.setShrinkToFit as any).call({}, true));
        assert.strictEqual(format.setShrinkToFit(true), format);
        assert.strictEqual(format.shrinkToFit(), true);
        assert.strictEqual(format.setShrinkToFit(false).shrinkToFit(), false);
    });

    it('format.setBorder sets all borders', () => {
        const expectAllBordersToBe = (type: number) => {
            ['Left', 'Right', 'Bottom', 'Top'].forEach((border) => {
                assert.strictEqual((format as any)['border' + border](), type);
            });
        };

        assert.throws(() => (format.setBorder as any).format.call(undefined, 'a'));
        assert.throws(() => (format.setBorder as any).call({}, xl.BORDERSTYLE_THIN));
        assert.strictEqual(format.setBorder(xl.BORDERSTYLE_THIN), format);
        expectAllBordersToBe(xl.BORDERSTYLE_THIN);
        assert.strictEqual(format.setBorder(xl.BORDERSTYLE_MEDIUM), format);
        expectAllBordersToBe(xl.BORDERSTYLE_MEDIUM);
    });

    it('format.setBorderColor sets the color of all borders', () => {
        const expectAllBordersToBe = (color: number) => {
            ['Left', 'Right', 'Bottom', 'Top'].forEach((border) => {
                assert.strictEqual((format as any)['border' + border + 'Color'](), color);
            });
        };

        assert.throws(() => (format.setBorderColor as any).call(format, 'a'));
        assert.throws(() => (format.setBorderColor as any).call({}, xl.COLOR_YELLOW));
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
            assert.throws(() => (format as any)[styleGetter].call({}));
            (format as any)[styleGetter]();
        });

        it(util.format('format.%s sets the %s border style', styleSetter, border.toLowerCase()), () => {
            assert.throws(() => (format as any)[styleSetter].call(format, 'a'));
            assert.throws(() => (format as any)[styleSetter].call({}, xl.BORDERSTYLE_THIN));
            assert.strictEqual((format as any)[styleSetter](xl.BORDERSTYLE_THIN), format);
            assert.strictEqual((format as any)[styleGetter](), xl.BORDERSTYLE_THIN);
            assert.strictEqual(
                (format as any)[styleSetter](xl.BORDERSTYLE_MEDIUM)[styleGetter](),
                xl.BORDERSTYLE_MEDIUM,
            );
        });
    });

    ['Left', 'Right', 'Top', 'Bottom', 'Diagonal'].forEach((border) => {
        const colorGetter = 'border' + border + 'Color',
            colorSetter = 'setBorder' + border + 'Color';

        it(util.format('format.%s returns the %s border color', colorGetter, border.toLowerCase()), () => {
            assert.throws(() => (format as any)[colorGetter]().call({}));
            (format as any)[colorGetter]();
        });

        it(util.format('format.%s sets the %s border color', colorSetter, border.toLowerCase()), () => {
            assert.throws(() => (format as any)[colorSetter].call(format, 'a'));
            assert.throws(() => (format as any)[colorSetter].call({}, xl.COLOR_YELLOW));
            assert.strictEqual((format as any)[colorSetter](xl.COLOR_YELLOW), format);
            assert.strictEqual((format as any)[colorGetter](), xl.COLOR_YELLOW);
            assert.strictEqual((format as any)[colorSetter](xl.COLOR_GREEN)[colorGetter](), xl.COLOR_GREEN);
        });
    });

    it('format.borderDiagonal returns diagonal border style', () => {
        assert.throws(() => (format.borderDiagonal as any).call({}));
        format.borderDiagonal();
    });

    it('format.setBorderDiagonal sets diagonal border style', () => {
        assert.throws(() => (format.setBorderDiagonal as any).call(format, 'a'));
        assert.throws(() => (format.setBorderDiagonal as any).call({}, xl.BORDERDIAGONAL_DOWN));
        assert.strictEqual(format.setBorderDiagonal(xl.BORDERDIAGONAL_DOWN), format);
        assert.strictEqual(format.borderDiagonal(), xl.BORDERDIAGONAL_DOWN);
        assert.strictEqual(format.setBorderDiagonal(xl.BORDERDIAGONAL_UP).borderDiagonal(), xl.BORDERDIAGONAL_UP);
    });

    it('format.fillPattern returns the fill pattern', () => {
        assert.throws(() => (format.fillPattern as any)().call({}));
        format.fillPattern();
    });

    it('format.setFillPattern sets the fill pattern', () => {
        assert.throws(() => (format.setFillPattern as any).call(format, 'a'));
        assert.throws(() => (format.setFillPattern as any).call({}, xl.FILLPATTERN_GRAY50));
        assert.strictEqual(format.setFillPattern(xl.FILLPATTERN_GRAY50), format);
        assert.strictEqual(format.fillPattern(), xl.FILLPATTERN_GRAY50);
        assert.strictEqual(format.setFillPattern(xl.FILLPATTERN_GRAY75).fillPattern(), xl.FILLPATTERN_GRAY75);
    });

    ['Foreground', 'Background'].forEach((layer) => {
        const setter = 'setPattern' + layer + 'Color',
            getter = 'pattern' + layer + 'Color';

        it(util.format('format.%s returns %s pattern color', getter, layer.toLowerCase()), () => {
            assert.throws(() => (format as any)[getter].call({}));
            (format as any)[getter]();
        });

        it(util.format('format.%s sets %s pattern color', setter, layer.toLowerCase()), () => {
            assert.throws(() => (format as any)[setter].call(format, 'a'));
            assert.throws(() => (format as any)[setter].call({}, xl.COLOR_YELLOW));
            assert.strictEqual((format as any)[setter](xl.COLOR_YELLOW), format);
            assert.strictEqual((format as any)[getter](), xl.COLOR_YELLOW);
            assert.strictEqual((format as any)[setter](xl.COLOR_GREEN)[getter](), xl.COLOR_GREEN);
        });
    });

    it('format.locked checks whether the format is locked', () => {
        assert.throws(() => (format.locked as any).call({}));
        format.locked();
    });

    it('format.setLocked toggles a format locked', () => {
        assert.throws(() => (format.setLocked as any).call(format, 1));
        assert.throws(() => (format.setLocked as any).call({}, true));
        assert.strictEqual(format.setLocked(true), format);
        assert.strictEqual(format.locked(), true);
        assert.strictEqual(format.setLocked(false).locked(), false);
        assert.strictEqual(format.setLocked().locked(), true);
    });

    it('format.hidden checks whether the format is hidden', () => {
        assert.throws(() => (format.hidden as any).call({}));
        format.hidden();
    });

    it('format.setHidden toggles a format hidden', () => {
        assert.throws(() => (format.setHidden as any).call(format, 1));
        assert.throws(() => (format.setHidden as any).call({}, true));
        assert.strictEqual(format.setHidden(true), format);
        assert.strictEqual(format.hidden(), true);
        assert.strictEqual(format.setHidden(false).hidden(), false);
        assert.strictEqual(format.setHidden().hidden(), true);
    });
});
