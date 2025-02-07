/**
 * The MIT License (MIT)
 *
 * Copyright (c) 2013 Christian Speckner <cnspeckn@googlemail.com>
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

#ifndef BINDINGS_SHEET_H
#define BINDINGS_SHEET_H

#include "book_wrapper.h"
#include "common.h"
#include "wrapper.h"

namespace node_libxl {

    class Sheet : public Wrapper<libxl::Sheet>,
                  public BookWrapper

    {
       public:
        Sheet(libxl::Sheet* sheet, v8::Local<v8::Value> book);

        static void Initialize(v8::Local<v8::Object> exports);

        static Sheet* Unwrap(v8::Local<v8::Value> object) {
            return Wrapper<libxl::Sheet>::Unwrap<Sheet>(object);
        }

        static v8::Local<v8::Object> NewInstance(libxl::Sheet* sheet, v8::Local<v8::Value> book);

       protected:
        static NAN_METHOD(CellType);
        static NAN_METHOD(IsFormula);
        static NAN_METHOD(CellFormat);
        static NAN_METHOD(SetCellFormat);
        static NAN_METHOD(ReadStr);
        static NAN_METHOD(WriteStr);
        static NAN_METHOD(ReadNum);
        static NAN_METHOD(WriteNum);
        static NAN_METHOD(ReadBool);
        static NAN_METHOD(WriteBool);
        static NAN_METHOD(ReadBlank);
        static NAN_METHOD(WriteBlank);
        static NAN_METHOD(ReadFormula);
        static NAN_METHOD(WriteFormula);
        static NAN_METHOD(WriteFormulaNum);
        static NAN_METHOD(WriteFormulaStr);
        static NAN_METHOD(WriteFormulaBool);
        static NAN_METHOD(ReadComment);
        static NAN_METHOD(WriteComment);
        static NAN_METHOD(RemoveComment);
        static NAN_METHOD(IsDate);
        static NAN_METHOD(IsRichStr);
        static NAN_METHOD(ReadError);
        static NAN_METHOD(WriteError);
        static NAN_METHOD(ColWidth);
        static NAN_METHOD(RowHeight);
        static NAN_METHOD(ColFormat);
        static NAN_METHOD(RowFormat);
        static NAN_METHOD(ColWidthPx);
        static NAN_METHOD(RowHeightPx);
        static NAN_METHOD(SetCol);
        static NAN_METHOD(SetRow);
        static NAN_METHOD(SetColPx);
        static NAN_METHOD(SetRowPx);
        static NAN_METHOD(RowHidden);
        static NAN_METHOD(SetRowHidden);
        static NAN_METHOD(ColHidden);
        static NAN_METHOD(SetColHidden);
        static NAN_METHOD(DefaultRowHeight);
        static NAN_METHOD(SetDefaultRowHeight);
        static NAN_METHOD(GetMerge);
        static NAN_METHOD(SetMerge);
        static NAN_METHOD(DelMerge);
        static NAN_METHOD(MergeSize);
        static NAN_METHOD(Merge);
        static NAN_METHOD(DelMergeByIndex);
        static NAN_METHOD(PictureSize);
        static NAN_METHOD(GetPicture);
        static NAN_METHOD(SetPicture);
        static NAN_METHOD(SetPicture2);
        static NAN_METHOD(RemovePicture);
        static NAN_METHOD(RemovePictureByIndex);
        static NAN_METHOD(GetHorPageBreak);
        static NAN_METHOD(GetHorPageBreakSize);
        static NAN_METHOD(GetVerPageBreak);
        static NAN_METHOD(GetVerPageBreakSize);
        static NAN_METHOD(SetHorPageBreak);
        static NAN_METHOD(SetVerPageBreak);
        static NAN_METHOD(Split);
        static NAN_METHOD(SplitInfo);
        static NAN_METHOD(GroupRows);
        static NAN_METHOD(GroupCols);
        static NAN_METHOD(GroupSummaryBelow);
        static NAN_METHOD(SetGroupSummaryBelow);
        static NAN_METHOD(GroupSummaryRight);
        static NAN_METHOD(SetGroupSummaryRight);
        static NAN_METHOD(Clear);
        static NAN_METHOD(InsertRow);
        static NAN_METHOD(InsertRowAsync);
        static NAN_METHOD(InsertCol);
        static NAN_METHOD(InsertColAsync);
        static NAN_METHOD(RemoveRow);
        static NAN_METHOD(RemoveRowAsync);
        static NAN_METHOD(RemoveCol);
        static NAN_METHOD(RemoveColAsync);
        static NAN_METHOD(CopyCell);
        static NAN_METHOD(FirstRow);
        static NAN_METHOD(LastRow);
        static NAN_METHOD(FirstCol);
        static NAN_METHOD(LastCol);
        static NAN_METHOD(FirstFilledRow);
        static NAN_METHOD(LastFilledRow);
        static NAN_METHOD(FirstFilledCol);
        static NAN_METHOD(LastFilledCol);
        static NAN_METHOD(DisplayGridlines);
        static NAN_METHOD(SetDisplayGridlines);
        static NAN_METHOD(PrintGridlines);
        static NAN_METHOD(SetPrintGridlines);
        static NAN_METHOD(Zoom);
        static NAN_METHOD(SetZoom);
        static NAN_METHOD(PrintZoom);
        static NAN_METHOD(SetPrintZoom);
        static NAN_METHOD(GetPrintFit);
        static NAN_METHOD(SetPrintFit);
        static NAN_METHOD(Landscape);
        static NAN_METHOD(SetLandscape);
        static NAN_METHOD(Paper);
        static NAN_METHOD(SetPaper);
        static NAN_METHOD(Header);
        static NAN_METHOD(SetHeader);
        static NAN_METHOD(HeaderMargin);
        static NAN_METHOD(Footer);
        static NAN_METHOD(SetFooter);
        static NAN_METHOD(FooterMargin);
        static NAN_METHOD(HCenter);
        static NAN_METHOD(SetHCenter);
        static NAN_METHOD(VCenter);
        static NAN_METHOD(SetVCenter);
        static NAN_METHOD(MarginLeft);
        static NAN_METHOD(SetMarginLeft);
        static NAN_METHOD(MarginRight);
        static NAN_METHOD(SetMarginRight);
        static NAN_METHOD(MarginTop);
        static NAN_METHOD(SetMarginTop);
        static NAN_METHOD(MarginBottom);
        static NAN_METHOD(SetMarginBottom);
        static NAN_METHOD(PrintRowCol);
        static NAN_METHOD(SetPrintRowCol);
        static NAN_METHOD(PrintRepeatRows);
        static NAN_METHOD(SetPrintRepeatRows);
        static NAN_METHOD(PrintRepeatCols);
        static NAN_METHOD(SetPrintRepeatCols);
        static NAN_METHOD(SetPrintArea);
        static NAN_METHOD(ClearPrintRepeats);
        static NAN_METHOD(ClearPrintArea);
        static NAN_METHOD(GetNamedRange);
        static NAN_METHOD(SetNamedRange);
        static NAN_METHOD(DelNamedRange);
        static NAN_METHOD(NamedRangeSize);
        static NAN_METHOD(NamedRange);
        static NAN_METHOD(GetTable);
        static NAN_METHOD(TableSize);
        static NAN_METHOD(Table);
        static NAN_METHOD(HyperlinkSize);
        static NAN_METHOD(Hyperlink);
        static NAN_METHOD(DelHyperlink);
        static NAN_METHOD(AddHyperlink);
        static NAN_METHOD(HyperlinkIndex);
        static NAN_METHOD(Name);
        static NAN_METHOD(SetName);
        static NAN_METHOD(Protect);
        static NAN_METHOD(SetProtect);
        static NAN_METHOD(RightToLeft);
        static NAN_METHOD(SetRightToLeft);
        static NAN_METHOD(Hidden);
        static NAN_METHOD(SetHidden);
        static NAN_METHOD(GetTopLeftView);
        static NAN_METHOD(SetTopLeftView);
        static NAN_METHOD(AddrToRowCol);
        static NAN_METHOD(RowColToAddr);
        static NAN_METHOD(SetAutoFitArea);
        static NAN_METHOD(TabColor);
        static NAN_METHOD(SetTabColor);
        static NAN_METHOD(SetTabColorComponents);
        static NAN_METHOD(GetTabColor);
        static NAN_METHOD(AddIgnoredError);

       private:
        const libxl::Sheet* wrappedSheet;

       private:
        Sheet(const Sheet&);
        const Sheet& operator=(const Sheet&);
    };

}  // namespace node_libxl

#endif  // BINDINGS_SHEET_H
