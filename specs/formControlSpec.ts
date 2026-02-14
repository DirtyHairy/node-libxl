import { describe, it, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'fs';
import * as xl from '../lib/libxl';
import { getXlsmFormControlFile } from './testUtils';

let fixture = fs.readFileSync(getXlsmFormControlFile());

describe('FormControl', () => {
    let book: xl.Book, sheet: xl.Sheet, formControl: xl.FormControl;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.loadRawSync(fixture);

        sheet = book.getSheet(0);
        formControl = sheet.formControl(0);
    });

    it('objectType returns the type of the form control', () => {
        assert.throws(() => (formControl.objectType as any).call({}));

        assert.strictEqual(formControl.objectType(), xl.OBJECT_LIST);
    });

    it('checked and setChecked control the cell reference in a group box is linked to', () => {
        assert.throws(() => (formControl.checked as any).call({}));
        assert.throws(() => (formControl.setChecked as any).call(formControl, 'a'));
        assert.throws(() => (formControl.setChecked as any).call(formControl, xl.CHECKED_TYPE_CHECKED));

        assert.strictEqual(formControl.setChecked(xl.CHECKEDTYPE_CHECKED), formControl);
        assert.strictEqual(formControl.checked(), xl.CHECKEDTYPE_CHECKED);
    });

    it('fmlaGroup and setFmlaGroup control the cell reference with range of source data cells', () => {
        assert.throws(() => (formControl.fmlaGroup as any).call({}));
        assert.throws(() => (formControl.setFmlaGroup as any).call(formControl, 1));
        assert.throws(() => (formControl.setFmlaGroup as any).call({}, 'a'));

        assert.strictEqual(formControl.setFmlaGroup('a'), formControl);
        assert.strictEqual(formControl.fmlaGroup(), 'a');
    });

    it('fmlaLink and setFmlaLink control the  cell reference is linked to', () => {
        assert.throws(() => (formControl.fmlaLink as any).call({}));
        assert.throws(() => (formControl.setFmlaLink as any).call(formControl, 1));
        assert.throws(() => (formControl.setFmlaLink as any).call({}, 'a'));

        assert.strictEqual(formControl.setFmlaLink('a'), formControl);
        assert.strictEqual(formControl.fmlaLink(), 'a');
    });

    it('fmlaRange and setFmlaRange control the formula range of the control', () => {
        assert.throws(() => (formControl.fmlaRange as any).call({}));
        assert.throws(() => (formControl.setFmlaRange as any).call(formControl, 1));
        assert.throws(() => (formControl.setFmlaRange as any).call({}, 'a'));

        assert.strictEqual(formControl.setFmlaRange('a'), formControl);
        assert.strictEqual(formControl.fmlaRange(), 'a');
    });

    it("fmlaTxbx and setFmlaTxbx control the cell reference with the source data that the form control object's data is linked to", () => {
        assert.throws(() => (formControl.fmlaTxbx as any).call({}));
        assert.throws(() => (formControl.setFmlaTxbx as any).call(formControl, 1));
        assert.throws(() => (formControl.setFmlaTxbx as any).call({}, 'a'));

        assert.strictEqual(formControl.setFmlaTxbx('a'), formControl);
        assert.strictEqual(formControl.fmlaTxbx(), 'a');
    });

    it('name return the name of the control', () => {
        assert.throws(() => (formControl.name as any).call({}));

        assert.strictEqual(formControl.name(), 'listfield');
    });

    it("linkedCell returns the worksheet range linked to the control's value", () => {
        assert.throws(() => (formControl.linkedCell as any).call({}));

        assert.strictEqual(formControl.linkedCell(), '');
    });

    it('listFillRange returns the range of source data cells used to populate the list box', () => {
        assert.throws(() => (formControl.listFillRange as any).call({}));

        assert.strictEqual(formControl.listFillRange(), '');
    });

    it('macro returns the macro function associated with the control', () => {
        assert.throws(() => (formControl.macro as any).call({}));

        assert.strictEqual(formControl.macro(), '[0]!Macro1');
    });

    it('altText returns the alternate text associated with the control', () => {
        assert.throws(() => (formControl.altText as any).call({}));

        assert.strictEqual(formControl.altText(), 'Alternate');
    });

    it('locked checks whether the control is locked', () => {
        assert.throws(() => (formControl.locked as any).call({}));

        assert.strictEqual(formControl.locked(), true);
    });

    it('defaultSize checks whether the control is at its default size', () => {
        assert.throws(() => (formControl.defaultSize as any).call({}));

        assert.strictEqual(formControl.defaultSize(), false);
    });

    it('print checks whether the control should be printed', () => {
        assert.throws(() => (formControl.print as any).call({}));

        assert.strictEqual(formControl.print(), true);
    });

    it('disabled checks whether the control is disabled', () => {
        assert.throws(() => (formControl.disabled as any).call({}));

        assert.strictEqual(formControl.disabled(), false);
    });

    it('the items family of function manages the items in a list control', () => {
        assert.throws(() => (formControl.item as any).call(formControl, 'a'));
        assert.throws(() => (formControl.item as any).call({}, 1));

        assert.throws(() => (formControl.itemSize as any).call({}));
        assert.throws(() => (formControl.clearItems as any).call({}));

        assert.throws(() => (formControl.addItem as any).call(formControl, 1));
        assert.throws(() => (formControl.addItem as any).call({}, 'a'));

        assert.throws(() => (formControl.insertItem as any).call(formControl, 1, 1));
        assert.throws(() => (formControl.insertItem as any).call({}, 1, 'a'));

        assert.strictEqual(formControl.clearItems(), formControl);

        assert.strictEqual(formControl.addItem('a'), formControl);
        assert.strictEqual(formControl.insertItem(0, 'b'), formControl);

        assert.strictEqual(formControl.itemSize(), 2);

        assert.strictEqual(formControl.item(0), 'b');
        assert.strictEqual(formControl.item(1), 'a');
    });

    it('dropLines and setDropLines control the number of lines in the drop-down before scroll bars are added', () => {
        assert.throws(() => (formControl.dropLines as any).call({}));
        assert.throws(() => (formControl.setDropLines as any).call(formControl, 'a'));
        assert.throws(() => (formControl.setDropLines as any).call({}, 1));

        assert.strictEqual(formControl.setDropLines(1), formControl);
        assert.strictEqual(formControl.dropLines(), 1);
    });

    it('dx and setDx control the with of the scroll bar', () => {
        assert.throws(() => (formControl.dx as any).call({}));
        assert.throws(() => (formControl.setDx as any).call(formControl, 'a'));
        assert.throws(() => (formControl.setDx as any).call({}, 1));

        assert.strictEqual(formControl.setDx(1), formControl);
        assert.strictEqual(formControl.dx(), 1);
    });

    it('firstButton and setFirstButton control whether the control is the first button in a list of radios', () => {
        assert.throws(() => (formControl.firstButton as any).call({}));
        assert.throws(() => (formControl.setFirstButton as any).call(formControl, 1));
        assert.throws(() => (formControl.setFirstButton as any).call({}, false));

        assert.strictEqual(formControl.setFirstButton(false), formControl);
        assert.strictEqual(formControl.firstButton(), false);
    });

    it('horiz and setHoriz control the direction of the scroll bar', () => {
        assert.throws(() => (formControl.horiz as any).call({}));
        assert.throws(() => (formControl.setHoriz as any).call(formControl, 1));
        assert.throws(() => (formControl.setHoriz as any).call({}, false));

        assert.strictEqual(formControl.setHoriz(false), formControl);
        assert.strictEqual(formControl.horiz(), false);
    });

    it('inc and setInc control the increment of the control', () => {
        assert.throws(() => (formControl.inc as any).call({}));
        assert.throws(() => (formControl.setInc as any).call(formControl, 'a'));
        assert.throws(() => (formControl.setInc as any).call({}, 1));

        assert.strictEqual(formControl.setInc(1), formControl);
        assert.strictEqual(formControl.inc(), 1);
    });

    it('getMin and getMin control the minimum value of the control', () => {
        assert.throws(() => (formControl.getMin as any).call({}));
        assert.throws(() => (formControl.setMin as any).call(formControl, 'a'));
        assert.throws(() => (formControl.setMin as any).call({}, 1));

        assert.strictEqual(formControl.setMin(1), formControl);
        assert.strictEqual(formControl.getMin(), 1);
    });

    it('getMax and getMax control the maximum value of the control', () => {
        assert.throws(() => (formControl.getMax as any).call({}));
        assert.throws(() => (formControl.setMax as any).call(formControl, 'a'));
        assert.throws(() => (formControl.setMax as any).call({}, 1));

        assert.strictEqual(formControl.setMax(1), formControl);
        assert.strictEqual(formControl.getMax(), 1);
    });

    it('multiSel and setMultiSel control the list of selected cells', () => {
        assert.throws(() => (formControl.multiSel as any).call({}));
        assert.throws(() => (formControl.setMultiSel as any).call(formControl, 1));
        assert.throws(() => (formControl.setMultiSel as any).call({}, 'a1,a2'));

        assert.strictEqual(formControl.setMultiSel('a1,a2'), formControl);
        assert.strictEqual(formControl.multiSel(), 'a1,a2');
    });

    it('sel and setSel control the index of the selected item', () => {
        assert.throws(() => (formControl.sel as any).call({}));
        assert.throws(() => (formControl.setSel as any).call(formControl, 'a'));
        assert.throws(() => (formControl.setSel as any).call({}, 1));

        assert.strictEqual(formControl.setSel(1), formControl);
        assert.strictEqual(formControl.sel(), 1);
    });

    it('fromAnchor returns the top left position of the control', () => {
        assert.throws(() => (formControl.fromAnchor as any).call({}));

        assert.deepStrictEqual(formControl.fromAnchor(), { row: 1, col: 1, colOff: 127000, rowOff: 76200 });
    });

    it('toAnchor returns the top left position of the control', () => {
        assert.throws(() => (formControl.toAnchor as any).call({}));

        assert.deepStrictEqual(formControl.toAnchor(), { row: 4, col: 5, colOff: 406400, rowOff: 165100 });
    });
});
