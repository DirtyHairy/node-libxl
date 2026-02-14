const exp = require('constants');

const { describe, it, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
var xl = require('../lib/libxl'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow,
    fs = require('fs');

let fixture = fs.readFileSync(testUtils.getXlsmFormControlFile());

describe('FormControl', () => {
    let book, sheet, formControl;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.loadRawSync(fixture);

        sheet = book.getSheet(0);
        formControl = sheet.formControl(0);
    });

    it('objectType returns the type of the form control', () => {
        shouldThrow(formControl.objectType, {});

        assert.strictEqual(formControl.objectType(), xl.OBJECT_LIST);
    });

    it('checked and setChecked control the cell reference in a group box is linked to', () => {
        shouldThrow(formControl.checked, {});
        shouldThrow(formControl.setChecked, formControl, 'a');
        shouldThrow(formControl.setChecked, formControl, xl.CHECKED_TYPE_CHECKED);

        assert.strictEqual(formControl.setChecked(xl.CHECKEDTYPE_CHECKED), formControl);
        assert.strictEqual(formControl.checked(), xl.CHECKEDTYPE_CHECKED);
    });

    it('fmlaGroup and setFmlaGroup control the cell reference with range of source data cells', () => {
        shouldThrow(formControl.fmlaGroup, {});
        shouldThrow(formControl.setFmlaGroup, formControl, 1);
        shouldThrow(formControl.setFmlaGroup, {}, 'a');

        assert.strictEqual(formControl.setFmlaGroup('a'), formControl);
        assert.strictEqual(formControl.fmlaGroup(), 'a');
    });

    it('fmlaLink and setFmlaLink control the  cell reference is linked to', () => {
        shouldThrow(formControl.fmlaLink, {});
        shouldThrow(formControl.setFmlaLink, formControl, 1);
        shouldThrow(formControl.setFmlaLink, {}, 'a');

        assert.strictEqual(formControl.setFmlaLink('a'), formControl);
        assert.strictEqual(formControl.fmlaLink(), 'a');
    });

    it('fmlaRange and setFmlaRange control the formula range of the control', () => {
        shouldThrow(formControl.fmlaRange, {});
        shouldThrow(formControl.setFmlaRange, formControl, 1);
        shouldThrow(formControl.setFmlaRange, {}, 'a');

        assert.strictEqual(formControl.setFmlaRange('a'), formControl);
        assert.strictEqual(formControl.fmlaRange(), 'a');
    });

    it("fmlaTxbx and setFmlaTxbx control the cell reference with the source data that the form control object's data is linked to", () => {
        shouldThrow(formControl.fmlaTxbx, {});
        shouldThrow(formControl.setFmlaTxbx, formControl, 1);
        shouldThrow(formControl.setFmlaTxbx, {}, 'a');

        assert.strictEqual(formControl.setFmlaTxbx('a'), formControl);
        assert.strictEqual(formControl.fmlaTxbx(), 'a');
    });

    it('name return the name of the control', () => {
        shouldThrow(formControl.name, {});

        assert.strictEqual(formControl.name(), 'listfield');
    });

    it("linkedCell returns the worksheet range linked to the control's value", () => {
        shouldThrow(formControl.linkedCell, {});

        assert.strictEqual(formControl.linkedCell(), '');
    });

    it('listFillRange returns the range of source data cells used to populate the list box', () => {
        shouldThrow(formControl.listFillRange, {});

        assert.strictEqual(formControl.listFillRange(), '');
    });

    it('macro returns the macro function associated with the control', () => {
        shouldThrow(formControl.macro, {});

        assert.strictEqual(formControl.macro(), '[0]!Macro1');
    });

    it('altText returns the alternate text associated with the control', () => {
        shouldThrow(formControl.altText, {});

        assert.strictEqual(formControl.altText(), 'Alternate');
    });

    it('locked checks whether the control is locked', () => {
        shouldThrow(formControl.locked, {});

        assert.strictEqual(formControl.locked(), true);
    });

    it('defaultSize checks whether the control is at its default size', () => {
        shouldThrow(formControl.defaultSize, {});

        assert.strictEqual(formControl.defaultSize(), false);
    });

    it('print checks whether the control should be printed', () => {
        shouldThrow(formControl.print, {});

        assert.strictEqual(formControl.print(), true);
    });

    it('disabled checks whether the control is disabled', () => {
        shouldThrow(formControl.disabled, {});

        assert.strictEqual(formControl.disabled(), false);
    });

    it('the items family of function manages the items in a list control', () => {
        shouldThrow(formControl.item, formControl, 'a');
        shouldThrow(formControl.item, {}, 1);

        shouldThrow(formControl.itemSize, {});
        shouldThrow(formControl.clearItems, {});

        shouldThrow(formControl.addItem, formControl, 1);
        shouldThrow(formControl.addItem, {}, 'a');

        shouldThrow(formControl.insertItem, formControl, 1, 1);
        shouldThrow(formControl.insertItem, {}, 1, 'a');

        assert.strictEqual(formControl.clearItems(), formControl);

        assert.strictEqual(formControl.addItem('a'), formControl);
        assert.strictEqual(formControl.insertItem(0, 'b'), formControl);

        assert.strictEqual(formControl.itemSize(), 2);

        assert.strictEqual(formControl.item(0), 'b');
        assert.strictEqual(formControl.item(1), 'a');
    });

    it('dropLines and setDropLines control the number of lines in the drop-down before scroll bars are added', () => {
        shouldThrow(formControl.dropLines, {});
        shouldThrow(formControl.setDropLines, formControl, 'a');
        shouldThrow(formControl.setDropLines, {}, 1);

        assert.strictEqual(formControl.setDropLines(1), formControl);
        assert.strictEqual(formControl.dropLines(), 1);
    });

    it('dx and setDx control the with of the scroll bar', () => {
        shouldThrow(formControl.dx, {});
        shouldThrow(formControl.setDx, formControl, 'a');
        shouldThrow(formControl.setDx, {}, 1);

        assert.strictEqual(formControl.setDx(1), formControl);
        assert.strictEqual(formControl.dx(), 1);
    });

    it('firstButton and setFirstButton control whether the control is the first button in a list of radios', () => {
        shouldThrow(formControl.firstButton, {});
        shouldThrow(formControl.setFirstButton, formControl, 1);
        shouldThrow(formControl.setFirstButton, {}, false);

        assert.strictEqual(formControl.setFirstButton(false), formControl);
        assert.strictEqual(formControl.firstButton(), false);
    });

    it('horiz and setHoriz control the direction of the scroll bar', () => {
        shouldThrow(formControl.horiz, {});
        shouldThrow(formControl.setHoriz, formControl, 1);
        shouldThrow(formControl.setHoriz, {}, false);

        assert.strictEqual(formControl.setHoriz(false), formControl);
        assert.strictEqual(formControl.horiz(), false);
    });

    it('inc and setInc control the increment of the control', () => {
        shouldThrow(formControl.inc, {});
        shouldThrow(formControl.setInc, formControl, 'a');
        shouldThrow(formControl.setInc, {}, 1);

        assert.strictEqual(formControl.setInc(1), formControl);
        assert.strictEqual(formControl.inc(), 1);
    });

    it('getMin and getMin control the minimum value of the control', () => {
        shouldThrow(formControl.getMin, {});
        shouldThrow(formControl.setMin, formControl, 'a');
        shouldThrow(formControl.setMin, {}, 1);

        assert.strictEqual(formControl.setMin(1), formControl);
        assert.strictEqual(formControl.getMin(), 1);
    });

    it('getMax and getMax control the maximum value of the control', () => {
        shouldThrow(formControl.getMax, {});
        shouldThrow(formControl.setMax, formControl, 'a');
        shouldThrow(formControl.setMax, {}, 1);

        assert.strictEqual(formControl.setMax(1), formControl);
        assert.strictEqual(formControl.getMax(), 1);
    });

    it('multiSel and setMultiSel control the list of selected cells', () => {
        shouldThrow(formControl.multiSel, {});
        shouldThrow(formControl.setMultiSel, formControl, 1);
        shouldThrow(formControl.setMultiSel, {}, 'a1,a2');

        assert.strictEqual(formControl.setMultiSel('a1,a2'), formControl);
        assert.strictEqual(formControl.multiSel(), 'a1,a2');
    });

    it('sel and setSel control the index of the selected item', () => {
        shouldThrow(formControl.sel, {});
        shouldThrow(formControl.setSel, formControl, 'a');
        shouldThrow(formControl.setSel, {}, 1);

        assert.strictEqual(formControl.setSel(1), formControl);
        assert.strictEqual(formControl.sel(), 1);
    });

    it('fromAnchor returns the top left position of the control', () => {
        shouldThrow(formControl.fromAnchor, {});

        assert.deepStrictEqual(formControl.fromAnchor(), { row: 1, col: 1, colOff: 127000, rowOff: 76200 });
    });

    it('toAnchor returns the top left position of the control', () => {
        shouldThrow(formControl.toAnchor, {});

        assert.deepStrictEqual(formControl.toAnchor(), { row: 4, col: 5, colOff: 406400, rowOff: 165100 });
    });
});
