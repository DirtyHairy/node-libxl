import { describe, it, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'fs';
import xl from '../lib/libxl.js';
import { getXlsmFormControlFile } from './testUtils.js';

let fixture = fs.readFileSync(getXlsmFormControlFile());

describe('FormControl', () => {
    let book, sheet, formControl;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.loadRawSync(fixture);

        sheet = book.getSheet(0);
        formControl = sheet.formControl(0);
    });

    it('objectType returns the type of the form control', () => {
        assert.throws(() => formControl.objectType.call({}));

        assert.strictEqual(formControl.objectType(), xl.OBJECT_LIST);
    });

    it('checked and setChecked control the cell reference in a group box is linked to', () => {
        assert.throws(() => formControl.checked.call({}));
        assert.throws(() => formControl.setChecked.call(formControl, 'a'));
        assert.throws(() => formControl.setChecked.call(formControl, xl.CHECKED_TYPE_CHECKED));

        assert.strictEqual(formControl.setChecked(xl.CHECKEDTYPE_CHECKED), formControl);
        assert.strictEqual(formControl.checked(), xl.CHECKEDTYPE_CHECKED);
    });

    it('fmlaGroup and setFmlaGroup control the cell reference with range of source data cells', () => {
        assert.throws(() => formControl.fmlaGroup.call({}));
        assert.throws(() => formControl.setFmlaGroup.call(formControl, 1));
        assert.throws(() => formControl.setFmlaGroup.call({}, 'a'));

        assert.strictEqual(formControl.setFmlaGroup('a'), formControl);
        assert.strictEqual(formControl.fmlaGroup(), 'a');
    });

    it('fmlaLink and setFmlaLink control the  cell reference is linked to', () => {
        assert.throws(() => formControl.fmlaLink.call({}));
        assert.throws(() => formControl.setFmlaLink.call(formControl, 1));
        assert.throws(() => formControl.setFmlaLink.call({}, 'a'));

        assert.strictEqual(formControl.setFmlaLink('a'), formControl);
        assert.strictEqual(formControl.fmlaLink(), 'a');
    });

    it('fmlaRange and setFmlaRange control the formula range of the control', () => {
        assert.throws(() => formControl.fmlaRange.call({}));
        assert.throws(() => formControl.setFmlaRange.call(formControl, 1));
        assert.throws(() => formControl.setFmlaRange.call({}, 'a'));

        assert.strictEqual(formControl.setFmlaRange('a'), formControl);
        assert.strictEqual(formControl.fmlaRange(), 'a');
    });

    it("fmlaTxbx and setFmlaTxbx control the cell reference with the source data that the form control object's data is linked to", () => {
        assert.throws(() => formControl.fmlaTxbx.call({}));
        assert.throws(() => formControl.setFmlaTxbx.call(formControl, 1));
        assert.throws(() => formControl.setFmlaTxbx.call({}, 'a'));

        assert.strictEqual(formControl.setFmlaTxbx('a'), formControl);
        assert.strictEqual(formControl.fmlaTxbx(), 'a');
    });

    it('name return the name of the control', () => {
        assert.throws(() => formControl.name.call({}));

        assert.strictEqual(formControl.name(), 'listfield');
    });

    it("linkedCell returns the worksheet range linked to the control's value", () => {
        assert.throws(() => formControl.linkedCell.call({}));

        assert.strictEqual(formControl.linkedCell(), '');
    });

    it('listFillRange returns the range of source data cells used to populate the list box', () => {
        assert.throws(() => formControl.listFillRange.call({}));

        assert.strictEqual(formControl.listFillRange(), '');
    });

    it('macro returns the macro function associated with the control', () => {
        assert.throws(() => formControl.macro.call({}));

        assert.strictEqual(formControl.macro(), '[0]!Macro1');
    });

    it('altText returns the alternate text associated with the control', () => {
        assert.throws(() => formControl.altText.call({}));

        assert.strictEqual(formControl.altText(), 'Alternate');
    });

    it('locked checks whether the control is locked', () => {
        assert.throws(() => formControl.locked.call({}));

        assert.strictEqual(formControl.locked(), true);
    });

    it('defaultSize checks whether the control is at its default size', () => {
        assert.throws(() => formControl.defaultSize.call({}));

        assert.strictEqual(formControl.defaultSize(), false);
    });

    it('print checks whether the control should be printed', () => {
        assert.throws(() => formControl.print.call({}));

        assert.strictEqual(formControl.print(), true);
    });

    it('disabled checks whether the control is disabled', () => {
        assert.throws(() => formControl.disabled.call({}));

        assert.strictEqual(formControl.disabled(), false);
    });

    it('the items family of function manages the items in a list control', () => {
        assert.throws(() => formControl.item.call(formControl, 'a'));
        assert.throws(() => formControl.item.call({}, 1));

        assert.throws(() => formControl.itemSize.call({}));
        assert.throws(() => formControl.clearItems.call({}));

        assert.throws(() => formControl.addItem.call(formControl, 1));
        assert.throws(() => formControl.addItem.call({}, 'a'));

        assert.throws(() => formControl.insertItem.call(formControl, 1, 1));
        assert.throws(() => formControl.insertItem.call({}, 1, 'a'));

        assert.strictEqual(formControl.clearItems(), formControl);

        assert.strictEqual(formControl.addItem('a'), formControl);
        assert.strictEqual(formControl.insertItem(0, 'b'), formControl);

        assert.strictEqual(formControl.itemSize(), 2);

        assert.strictEqual(formControl.item(0), 'b');
        assert.strictEqual(formControl.item(1), 'a');
    });

    it('dropLines and setDropLines control the number of lines in the drop-down before scroll bars are added', () => {
        assert.throws(() => formControl.dropLines.call({}));
        assert.throws(() => formControl.setDropLines.call(formControl, 'a'));
        assert.throws(() => formControl.setDropLines.call({}, 1));

        assert.strictEqual(formControl.setDropLines(1), formControl);
        assert.strictEqual(formControl.dropLines(), 1);
    });

    it('dx and setDx control the with of the scroll bar', () => {
        assert.throws(() => formControl.dx.call({}));
        assert.throws(() => formControl.setDx.call(formControl, 'a'));
        assert.throws(() => formControl.setDx.call({}, 1));

        assert.strictEqual(formControl.setDx(1), formControl);
        assert.strictEqual(formControl.dx(), 1);
    });

    it('firstButton and setFirstButton control whether the control is the first button in a list of radios', () => {
        assert.throws(() => formControl.firstButton.call({}));
        assert.throws(() => formControl.setFirstButton.call(formControl, 1));
        assert.throws(() => formControl.setFirstButton.call({}, false));

        assert.strictEqual(formControl.setFirstButton(false), formControl);
        assert.strictEqual(formControl.firstButton(), false);
    });

    it('horiz and setHoriz control the direction of the scroll bar', () => {
        assert.throws(() => formControl.horiz.call({}));
        assert.throws(() => formControl.setHoriz.call(formControl, 1));
        assert.throws(() => formControl.setHoriz.call({}, false));

        assert.strictEqual(formControl.setHoriz(false), formControl);
        assert.strictEqual(formControl.horiz(), false);
    });

    it('inc and setInc control the increment of the control', () => {
        assert.throws(() => formControl.inc.call({}));
        assert.throws(() => formControl.setInc.call(formControl, 'a'));
        assert.throws(() => formControl.setInc.call({}, 1));

        assert.strictEqual(formControl.setInc(1), formControl);
        assert.strictEqual(formControl.inc(), 1);
    });

    it('getMin and getMin control the minimum value of the control', () => {
        assert.throws(() => formControl.getMin.call({}));
        assert.throws(() => formControl.setMin.call(formControl, 'a'));
        assert.throws(() => formControl.setMin.call({}, 1));

        assert.strictEqual(formControl.setMin(1), formControl);
        assert.strictEqual(formControl.getMin(), 1);
    });

    it('getMax and getMax control the maximum value of the control', () => {
        assert.throws(() => formControl.getMax.call({}));
        assert.throws(() => formControl.setMax.call(formControl, 'a'));
        assert.throws(() => formControl.setMax.call({}, 1));

        assert.strictEqual(formControl.setMax(1), formControl);
        assert.strictEqual(formControl.getMax(), 1);
    });

    it('multiSel and setMultiSel control the list of selected cells', () => {
        assert.throws(() => formControl.multiSel.call({}));
        assert.throws(() => formControl.setMultiSel.call(formControl, 1));
        assert.throws(() => formControl.setMultiSel.call({}, 'a1,a2'));

        assert.strictEqual(formControl.setMultiSel('a1,a2'), formControl);
        assert.strictEqual(formControl.multiSel(), 'a1,a2');
    });

    it('sel and setSel control the index of the selected item', () => {
        assert.throws(() => formControl.sel.call({}));
        assert.throws(() => formControl.setSel.call(formControl, 'a'));
        assert.throws(() => formControl.setSel.call({}, 1));

        assert.strictEqual(formControl.setSel(1), formControl);
        assert.strictEqual(formControl.sel(), 1);
    });

    it('fromAnchor returns the top left position of the control', () => {
        assert.throws(() => formControl.fromAnchor.call({}));

        assert.deepStrictEqual(formControl.fromAnchor(), { row: 1, col: 1, colOff: 127000, rowOff: 76200 });
    });

    it('toAnchor returns the top left position of the control', () => {
        assert.throws(() => formControl.toAnchor.call({}));

        assert.deepStrictEqual(formControl.toAnchor(), { row: 4, col: 5, colOff: 406400, rowOff: 165100 });
    });
});
