const exp = require('constants');

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

        expect(formControl.objectType()).toBe(xl.OBJECT_LIST);
    });

    it('checked and setChecked control the cell reference in a group box is linked to', () => {
        shouldThrow(formControl.checked, {});
        shouldThrow(formControl.setChecked, formControl, 'a');
        shouldThrow(formControl.setChecked, formControl, xl.CHECKED_TYPE_CHECKED);

        expect(formControl.setChecked(xl.CHECKEDTYPE_CHECKED)).toBe(formControl);
        expect(formControl.checked()).toBe(xl.CHECKEDTYPE_CHECKED);
    });

    it('fmlaGroup and setFmlaGroup control the cell reference with range of source data cells', () => {
        shouldThrow(formControl.fmlaGroup, {});
        shouldThrow(formControl.setFmlaGroup, formControl, 1);
        shouldThrow(formControl.setFmlaGroup, {}, 'a');

        expect(formControl.setFmlaGroup('a')).toBe(formControl);
        expect(formControl.fmlaGroup()).toBe('a');
    });

    it('fmlaLink and setFmlaLink control the  cell reference is linked to', () => {
        shouldThrow(formControl.fmlaLink, {});
        shouldThrow(formControl.setFmlaLink, formControl, 1);
        shouldThrow(formControl.setFmlaLink, {}, 'a');

        expect(formControl.setFmlaLink('a')).toBe(formControl);
        expect(formControl.fmlaLink()).toBe('a');
    });

    it('fmlaRange and setFmlaRange control the formula range of the control', () => {
        shouldThrow(formControl.fmlaRange, {});
        shouldThrow(formControl.setFmlaRange, formControl, 1);
        shouldThrow(formControl.setFmlaRange, {}, 'a');

        expect(formControl.setFmlaRange('a')).toBe(formControl);
        expect(formControl.fmlaRange()).toBe('a');
    });

    it("fmlaTxbx and setFmlaTxbx control the cell reference with the source data that the form control object's data is linked to", () => {
        shouldThrow(formControl.fmlaTxbx, {});
        shouldThrow(formControl.setFmlaTxbx, formControl, 1);
        shouldThrow(formControl.setFmlaTxbx, {}, 'a');

        expect(formControl.setFmlaTxbx('a')).toBe(formControl);
        expect(formControl.fmlaTxbx()).toBe('a');
    });

    it('name return the name of the control', () => {
        shouldThrow(formControl.name, {});

        expect(formControl.name()).toBe('listfield');
    });

    it("linkedCell returns the worksheet range linked to the control's value", () => {
        shouldThrow(formControl.linkedCell, {});

        expect(formControl.linkedCell()).toBe('');
    });

    it('listFillRange returns the range of source data cells used to populate the list box', () => {
        shouldThrow(formControl.listFillRange, {});

        expect(formControl.listFillRange()).toBe('');
    });

    it('macro returns the macro function associated with the control', () => {
        shouldThrow(formControl.macro, {});

        expect(formControl.macro()).toBe('[0]!Macro1');
    });

    it('altText returns the alternate text associated with the control', () => {
        shouldThrow(formControl.altText, {});

        expect(formControl.altText()).toBe('Alternate');
    });

    it('locked checks whether the control is locked', () => {
        shouldThrow(formControl.locked, {});

        expect(formControl.locked()).toBe(true);
    });

    it('defaultSize checks whether the control is at its default size', () => {
        shouldThrow(formControl.defaultSize, {});

        expect(formControl.defaultSize()).toBe(false);
    });

    it('print checks whether the control should be printed', () => {
        shouldThrow(formControl.print, {});

        expect(formControl.print()).toBe(true);
    });

    it('disabled checks whether the control is disabled', () => {
        shouldThrow(formControl.disabled, {});

        expect(formControl.disabled()).toBe(false);
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

        expect(formControl.clearItems()).toBe(formControl);

        expect(formControl.addItem('a')).toBe(formControl);
        expect(formControl.insertItem(0, 'b')).toBe(formControl);

        expect(formControl.itemSize()).toBe(2);

        expect(formControl.item(0)).toBe('b');
        expect(formControl.item(1)).toBe('a');
    });

    it('dropLines and setDropLines control the number of lines in the drop-down before scroll bars are added', () => {
        shouldThrow(formControl.dropLines, {});
        shouldThrow(formControl.setDropLines, formControl, 'a');
        shouldThrow(formControl.setDropLines, {}, 1);

        expect(formControl.setDropLines(1)).toBe(formControl);
        expect(formControl.dropLines()).toBe(1);
    });

    it('dx and setDx control the with of the scroll bar', () => {
        shouldThrow(formControl.dx, {});
        shouldThrow(formControl.setDx, formControl, 'a');
        shouldThrow(formControl.setDx, {}, 1);

        expect(formControl.setDx(1)).toBe(formControl);
        expect(formControl.dx()).toBe(1);
    });

    it('firstButton and setFirstButton control whether the control is the first button in a list of radios', () => {
        shouldThrow(formControl.firstButton, {});
        shouldThrow(formControl.setFirstButton, formControl, 1);
        shouldThrow(formControl.setFirstButton, {}, false);

        expect(formControl.setFirstButton(false)).toBe(formControl);
        expect(formControl.firstButton()).toBe(false);
    });

    it('horiz and setHoriz control the direction of the scroll bar', () => {
        shouldThrow(formControl.horiz, {});
        shouldThrow(formControl.setHoriz, formControl, 1);
        shouldThrow(formControl.setHoriz, {}, false);

        expect(formControl.setHoriz(false)).toBe(formControl);
        expect(formControl.horiz()).toBe(false);
    });

    it('inc and setInc control the increment of the control', () => {
        shouldThrow(formControl.inc, {});
        shouldThrow(formControl.setInc, formControl, 'a');
        shouldThrow(formControl.setInc, {}, 1);

        expect(formControl.setInc(1)).toBe(formControl);
        expect(formControl.inc()).toBe(1);
    });

    it('getMin and getMin control the minimum value of the control', () => {
        shouldThrow(formControl.getMin, {});
        shouldThrow(formControl.setMin, formControl, 'a');
        shouldThrow(formControl.setMin, {}, 1);

        expect(formControl.setMin(1)).toBe(formControl);
        expect(formControl.getMin()).toBe(1);
    });

    it('getMax and getMax control the maximum value of the control', () => {
        shouldThrow(formControl.getMax, {});
        shouldThrow(formControl.setMax, formControl, 'a');
        shouldThrow(formControl.setMax, {}, 1);

        expect(formControl.setMax(1)).toBe(formControl);
        expect(formControl.getMax()).toBe(1);
    });

    it('multiSel and setMultiSel control the list of selected cells', () => {
        shouldThrow(formControl.multiSel, {});
        shouldThrow(formControl.setMultiSel, formControl, 1);
        shouldThrow(formControl.setMultiSel, {}, 'a1,a2');

        expect(formControl.setMultiSel('a1,a2')).toBe(formControl);
        expect(formControl.multiSel()).toBe('a1,a2');
    });

    it('sel and setSel control the index of the selected item', () => {
        shouldThrow(formControl.sel, {});
        shouldThrow(formControl.setSel, formControl, 'a');
        shouldThrow(formControl.setSel, {}, 1);

        expect(formControl.setSel(1)).toBe(formControl);
        expect(formControl.sel()).toBe(1);
    });

    it('fromAnchor returns the top left position of the control', () => {
        shouldThrow(formControl.fromAnchor, {});

        expect(formControl.fromAnchor()).toEqual({ row: 1, col: 1, colOff: 127000, rowOff: 76200 });
    });

    it('toAnchor returns the top left position of the control', () => {
        shouldThrow(formControl.toAnchor, {});

        expect(formControl.toAnchor()).toEqual({ row: 4, col: 5, colOff: 406400, rowOff: 165100 });
    });
});
