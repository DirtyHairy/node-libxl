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

#include "sheet.h"

#include "assert.h"
#include "util.h"
#include "argument_helper.h"
#include "format.h"

using namespace v8;

namespace node_libxl {


// Lifecycle

Sheet::Sheet(libxl::Sheet* sheet, Handle<Value> book) :
    Wrapper<libxl::Sheet>(sheet),
    BookWrapper(book)
{}


Handle<Object> Sheet::NewInstance(
    libxl::Sheet* libxlSheet,
    Handle<Value> book)
{
    NanEscapableScope();

    Sheet* sheet = new Sheet(libxlSheet, book);

    Local<Object> that = NanNew(util::CallStubConstructor(
        NanNew(constructor)).As<Object>());

    sheet->Wrap(that);

    return NanEscapeScope(that);
}


// Wrappers


NAN_METHOD(Sheet::CellType) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    libxl::CellType cellType = that->GetWrapped()->cellType(row, col);
    if (cellType == libxl::CELLTYPE_ERROR) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(NanNew<Integer>(cellType));
}


NAN_METHOD(Sheet::IsFormula) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->isFormula(row, col)));
}


NAN_METHOD(Sheet::CellFormat) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    libxl::Format* libxlFormat = that->GetWrapped()->cellFormat(row, col);
    if (!libxlFormat) {
        return util::ThrowLibxlError(that);
    }

    
    NanReturnValue(Format::NewInstance(libxlFormat, that->GetBookHandle()));
}


NAN_METHOD(Sheet::SetCellFormat) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Format* format = arguments.GetWrapped<Format>(2);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);
    ASSERT_SAME_BOOK(that, format);

    that->GetWrapped()->setCellFormat(row, col, format->GetWrapped());

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::ReadStr) {
    NanScope();

    ArgumentHelper arguments(args);
    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Handle<Value> formatRef = args[2];
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    libxl::Format* libxlFormat = NULL;
    const char* value = that->GetWrapped()->readStr(row, col, &libxlFormat);
    if (!value) {
        return util::ThrowLibxlError(that);
    }

    if (formatRef->IsObject() && libxlFormat) {
        formatRef.As<Object>()->Set(NanNew<String>("format"),
            Format::NewInstance(libxlFormat, that->GetBookHandle()));
    }

    NanReturnValue(NanNew<String>(value));
}


NAN_METHOD(Sheet::WriteStr) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    String::Utf8Value value(arguments.GetString(2));
    Format* format = arguments.GetWrapped<Format>(3, NULL);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);
    if (format) {
        ASSERT_SAME_BOOK(that, format);
    }

    if (!that->GetWrapped()->
            writeStr(row, col, *value, format ? format->GetWrapped() : NULL))
    {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::ReadNum) {
    NanScope();

    ArgumentHelper arguments(args);
    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Handle<Value> formatRef = args[2];
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    libxl::Format* libxlFormat = NULL;
    double value = that->GetWrapped()->readNum(row, col, &libxlFormat);
    
    if (formatRef->IsObject() && libxlFormat) {
        formatRef.As<Object>()->Set(NanNew<String>("format"),
            Format::NewInstance(libxlFormat, that->GetBookHandle()));
    }

    NanReturnValue(NanNew<Number>(value));
}


NAN_METHOD(Sheet::WriteNum) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    double value = arguments.GetDouble(2);
    Format* format = arguments.GetWrapped<Format>(3, NULL);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);
    if (format) {
        ASSERT_SAME_BOOK(that, format);
    }

    if (!that->GetWrapped()->
            writeNum(row, col, value, format ? format->GetWrapped() : NULL))
    {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::ReadBool) {
    NanScope();

    ArgumentHelper arguments(args);
    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Handle<Value> formatRef = args[2];
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    libxl::Format* libxlFormat = NULL;
    bool value = that->GetWrapped()->readBool(row, col, &libxlFormat);
    
    if (formatRef->IsObject() && libxlFormat) {
        formatRef.As<Object>()->Set(NanNew<String>("format"),
            Format::NewInstance(libxlFormat, that->GetBookHandle()));
    }

    NanReturnValue(NanNew<Boolean>(value));
}


NAN_METHOD(Sheet::WriteBool) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    bool value = arguments.GetBoolean(2);
    Format* format = arguments.GetWrapped<Format>(3, NULL);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);
    if (format) {
        ASSERT_SAME_BOOK(that, format);
    }

    if (!that->GetWrapped()->
            writeBool(row, col, value, format ? format->GetWrapped() : NULL))
    {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::ReadBlank) {
    NanScope();

    ArgumentHelper arguments(args);
    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Handle<Value> formatRef = args[2];
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    libxl::Format* libxlFormat = NULL;
    if (!that->GetWrapped()->readBlank(row, col, &libxlFormat) || !libxlFormat) {
        return util::ThrowLibxlError(that);
    }

    Handle<Value> formatHandle = Format::NewInstance(libxlFormat,
        that->GetBookHandle());
    
    if (formatRef->IsObject() && libxlFormat) {
        formatRef.As<Object>()->Set(NanNew<String>("format"), formatHandle);
    }

    NanReturnValue(formatHandle);
}


NAN_METHOD(Sheet::WriteBlank) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Format* format = arguments.GetWrapped<Format>(2);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);
    ASSERT_SAME_BOOK(that, format);

    if (!that->GetWrapped()->
            writeBlank(row, col, format ? format->GetWrapped() : NULL))
    {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::ReadFormula) {
    NanScope();

    ArgumentHelper arguments(args);
    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Handle<Value> formatRef = args[2];
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    libxl::Format* libxlFormat = NULL;
    const char* value = that->GetWrapped()->readFormula(row, col, &libxlFormat);
    if (!value) {
        return util::ThrowLibxlError(that);
    }

    if (formatRef->IsObject() && libxlFormat) {
        formatRef.As<Object>()->Set(NanNew<String>("format"),
            Format::NewInstance(libxlFormat, that->GetBookHandle()));
    }

    NanReturnValue(NanNew<String>(value));
}


NAN_METHOD(Sheet::WriteFormula) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    String::Utf8Value value(arguments.GetString(2));
    Format* format = arguments.GetWrapped<Format>(3, NULL);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);
    if (format) {
        ASSERT_SAME_BOOK(that, format);
    }

    if (!that->GetWrapped()->
            writeFormula(row, col, *value, format? format->GetWrapped() : NULL))
    {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::ReadComment) {
    NanScope();

    ArgumentHelper arguments(args);
    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    const char* value = that->GetWrapped()->readComment(row, col);
    if (!value) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(NanNew<String>(value));
}


NAN_METHOD(Sheet::WriteComment) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    String::Utf8Value value(arguments.GetString(2));
    String::Utf8Value author(arguments.GetString(3, ""));
    int width = arguments.GetInt(4, 129);
    int height = arguments.GetInt(5, 75);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (args[3]->IsString()) {
        that->GetWrapped()->writeComment(row, col, *value, *author,
            width, height);
    } else {
        that->GetWrapped()->writeComment(row, col, *value, NULL,
            width, height);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::ReadError) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->readError(row, col)));
}


NAN_METHOD(Sheet::IsDate) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0),
        col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->isDate(row, col)));
}


NAN_METHOD(Sheet::ColWidth) {
    NanScope();

    ArgumentHelper arguments(args);

    int col = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Number>(that->GetWrapped()->colWidth(col)));
}


NAN_METHOD(Sheet::RowHeight) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Number>(that->GetWrapped()->rowHeight(row)));
}


NAN_METHOD(Sheet::SetCol) {
    NanScope();

    ArgumentHelper arguments(args);

    int first = arguments.GetInt(0);
    int last = arguments.GetInt(1);
    double width = arguments.GetDouble(2);
    Format* format = arguments.GetWrapped<Format>(3, NULL);
    bool hidden = arguments.GetBoolean(4, false);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);
    if (format) {
        ASSERT_SAME_BOOK(that, format);
    }

    if (!that->GetWrapped()->setCol(first, last, width,
            format ? format->GetWrapped() : NULL, hidden))
    {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::SetRow) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    double height = arguments.GetDouble(1);
    Format* format = arguments.GetWrapped<Format>(2, NULL);
    bool hidden = arguments.GetBoolean(3, false);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);
    if (format) {
        ASSERT_SAME_BOOK(that, format);
    }

    if (!that->GetWrapped()->setRow(row, height,
        format ? format->GetWrapped() : NULL, hidden))
    {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::RowHidden) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->rowHidden(row)));
}


NAN_METHOD(Sheet::SetRowHidden) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    bool hidden = arguments.GetBoolean(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setRowHidden(row, hidden)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::ColHidden) {
    NanScope();

    ArgumentHelper arguments(args);

    int col = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->colHidden(col)));
}


NAN_METHOD(Sheet::SetColHidden) {
    NanScope();

    ArgumentHelper arguments(args);

    int col = arguments.GetInt(0);
    bool hidden = arguments.GetBoolean(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setColHidden(col, hidden)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::GetMerge) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0),
        col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    int rowFirst, rowLast, colFirst, colLast;

    if (!that->GetWrapped()->getMerge(row, col, &rowFirst, &rowLast,
        &colFirst, &colLast))
    {
        return util::ThrowLibxlError(that);
    }

    Local<Object> result = NanNew<Object>();
    result->Set(NanNew<String>("rowFirst"), NanNew<Integer>(rowFirst));
    result->Set(NanNew<String>("rowLast"),  NanNew<Integer>(rowLast));
    result->Set(NanNew<String>("colFirst"), NanNew<Integer>(colFirst));
    result->Set(NanNew<String>("colLast"),  NanNew<Integer>(colLast));

    NanReturnValue(result);
}


NAN_METHOD(Sheet::SetMerge) {
    NanScope();

    ArgumentHelper arguments(args);

    int rowFirst = arguments.GetInt(0);
    int rowLast = arguments.GetInt(1);
    int colFirst = arguments.GetInt(2);
    int colLast = arguments.GetInt(3);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setMerge(rowFirst, rowLast, colFirst, colLast)) {
        return util::ThrowLibxlError(that);
    }
    
    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::DelMerge) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0),
        col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->delMerge(row, col)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::GetHorPageBreak) {
    NanScope();

    ArgumentHelper arguments(args);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->getHorPageBreak(index)));
}


NAN_METHOD(Sheet::GetHorPageBreakSize) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->getHorPageBreakSize()));
}


NAN_METHOD(Sheet::GetVerPageBreak) {
    NanScope();

    ArgumentHelper arguments(args);
    
    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->getVerPageBreak(index)));
}


NAN_METHOD(Sheet::GetVerPageBreakSize) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->getVerPageBreakSize()));
}


NAN_METHOD(Sheet::SetHorPageBreak) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0);
    bool pagebreak = arguments.GetBoolean(1, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setHorPageBreak(row, pagebreak)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::SetVerPageBreak) {
    NanScope();

    ArgumentHelper arguments(args);

    int col = arguments.GetInt(0);
    bool pagebreak = arguments.GetBoolean(1, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setVerPageBreak(col, pagebreak)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::Split) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0),
        col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->split(row, col);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::GroupRows) {
    NanScope();

    ArgumentHelper arguments(args);

    int rowFirst = arguments.GetInt(0),
        rowLast = arguments.GetInt(1);
    bool collapsed = arguments.GetBoolean(2, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->groupRows(rowFirst, rowLast, collapsed)) {
        return util::ThrowLibxlError(that);
    };

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::GroupCols) {
    NanScope();

    ArgumentHelper arguments(args);

    int colFirst = arguments.GetInt(0),
        colLast = arguments.GetInt(1);
    bool collapsed = arguments.GetBoolean(2, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->groupCols(colFirst, colLast, collapsed)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::GroupSummaryBelow) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->groupSummaryBelow()));
}


NAN_METHOD(Sheet::SetGroupSummaryBelow) {
    NanScope();

    ArgumentHelper arguments(args);

    bool summaryBelow = arguments.GetBoolean(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setGroupSummaryBelow(summaryBelow);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::GroupSummaryRight) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->groupSummaryRight()));
}


NAN_METHOD(Sheet::SetGroupSummaryRight) {
    NanScope();

    ArgumentHelper arguments(args);

    bool summaryRight = arguments.GetBoolean(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setGroupSummaryRight(summaryRight);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::Clear) {
    NanScope();

    ArgumentHelper arguments(args);

    int rowFirst    = arguments.GetInt(0, 0),
        rowLast     = arguments.GetInt(1, 65535),
        colFirst    = arguments.GetInt(2, 0),
        colLast     = arguments.GetInt(3, 255);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->clear(rowFirst, rowLast, colFirst, colLast);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::InsertRow) {
    NanScope();

    ArgumentHelper arguments(args);

    int rowFirst    = arguments.GetInt(0),
        rowLast     = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->insertRow(rowFirst, rowLast)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::InsertCol) {
    NanScope();

    ArgumentHelper arguments(args);

    int colFirst    = arguments.GetInt(0),
        colLast     = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->insertCol(colFirst, colLast)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::RemoveRow) {
    NanScope();

    ArgumentHelper arguments(args);

    int rowFirst    = arguments.GetInt(0),
        rowLast     = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->removeRow(rowFirst, rowLast)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::RemoveCol) {
    NanScope();

    ArgumentHelper arguments(args);

    int colFirst    = arguments.GetInt(0),
        colLast     = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->removeCol(colFirst, colLast)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::CopyCell) {
    NanScope();

    ArgumentHelper arguments(args);

    int rowSrc = arguments.GetInt(0),
        colSrc = arguments.GetInt(1),
        rowDst = arguments.GetInt(2),
        colDst = arguments.GetInt(3);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->copyCell(rowSrc, colSrc, rowDst, colDst)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::FirstRow) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->firstRow()));
}


NAN_METHOD(Sheet::LastRow) {
    NanScope();


    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->lastRow()));
}


NAN_METHOD(Sheet::FirstCol) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->firstCol()));
}

NAN_METHOD(Sheet::LastCol) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->lastCol()));
}


NAN_METHOD(Sheet::DisplayGridlines) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->displayGridlines()));
}


NAN_METHOD(Sheet::SetDisplayGridlines) {
    NanScope();

    ArgumentHelper arguments(args);

    bool displayGridlines = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);
    
    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setDisplayGridlines(displayGridlines);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::PrintGridlines) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->printGridlines()));
}


NAN_METHOD(Sheet::SetPrintGridlines) {
    NanScope();

    ArgumentHelper arguments(args);

    bool printGridlines = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintGridlines(printGridlines);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::Zoom) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->zoom()));
}


NAN_METHOD(Sheet::SetZoom) {
    NanScope();

    ArgumentHelper arguments(args);

    int zoom = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setZoom(zoom);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::PrintZoom) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->printZoom()));
}


NAN_METHOD(Sheet::SetPrintZoom) {
    NanScope();

    ArgumentHelper arguments(args);

    int zoom = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintZoom(zoom);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::GetPrintFit) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    int wPages, hPages;

    if (that->GetWrapped()->getPrintFit(&wPages, &hPages)) {
        Local<Object> result = NanNew<Object>();

        result->Set(NanNew<String>("wPages"),   NanNew<Integer>(wPages));
        result->Set(NanNew<String>("hPages"),   NanNew<Integer>(hPages));

        NanReturnValue(result);
    } else {
        NanReturnValue(NanFalse());
    }
}


NAN_METHOD(Sheet::SetPrintFit) {
    NanScope();

    ArgumentHelper arguments(args);

    int wPages = arguments.GetInt(0, 1),
        hPages = arguments.GetInt(1, 1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintFit(wPages, hPages);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::Landscape) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->landscape()));
}


NAN_METHOD(Sheet::SetLandscape) {
    NanScope();

    ArgumentHelper arguments(args);

    bool landscape = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setLandscape(landscape);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::Paper) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->paper()));
}


NAN_METHOD(Sheet::SetPaper) {
    NanScope();

    ArgumentHelper arguments(args);

    int paper = arguments.GetInt(0, libxl::PAPER_DEFAULT);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPaper(static_cast<libxl::Paper>(paper));

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::Header) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<String>(that->GetWrapped()->header()));
}


NAN_METHOD(Sheet::SetHeader) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value header(arguments.GetString(0));
    double margin = arguments.GetDouble(1, 0.5);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setHeader(*header, margin)) {
        util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::HeaderMargin) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Number>(that->GetWrapped()->headerMargin()));
}



NAN_METHOD(Sheet::Footer) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<String>(that->GetWrapped()->footer()));
}


NAN_METHOD(Sheet::SetFooter) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value footer(arguments.GetString(0));
    double margin = arguments.GetDouble(1, 0.5);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setFooter(*footer, margin)) {
        util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::FooterMargin) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Number>(that->GetWrapped()->footerMargin()));
}


NAN_METHOD(Sheet::HCenter) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->hCenter()));
}


NAN_METHOD(Sheet::SetHCenter) {
    NanScope();

    ArgumentHelper arguments(args);

    bool center = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setHCenter(center);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::VCenter) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->vCenter()));
}


NAN_METHOD(Sheet::SetVCenter) {
    NanScope();

    ArgumentHelper arguments(args);

    bool center = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setVCenter(center);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::MarginLeft) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Number>(that->GetWrapped()->marginLeft()));
}


NAN_METHOD(Sheet::SetMarginLeft) {
    NanScope();

    ArgumentHelper arguments(args);

    double margin = arguments.GetDouble(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setMarginLeft(margin);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::MarginRight) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Number>(that->GetWrapped()->marginRight()));
}


NAN_METHOD(Sheet::SetMarginRight) {
    NanScope();

    ArgumentHelper arguments(args);

    double margin = arguments.GetDouble(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setMarginRight(margin);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::MarginTop) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Number>(that->GetWrapped()->marginTop()));
}


NAN_METHOD(Sheet::SetMarginTop) {
    NanScope();

    ArgumentHelper arguments(args);

    double margin = arguments.GetDouble(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setMarginTop(margin);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::MarginBottom) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Number>(that->GetWrapped()->marginBottom()));
}


NAN_METHOD(Sheet::SetMarginBottom) {
    NanScope();

    ArgumentHelper arguments(args);

    double margin = arguments.GetDouble(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setMarginBottom(margin);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::PrintRowCol) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->printRowCol()));
}

NAN_METHOD(Sheet::SetPrintRowCol) {
    NanScope();

    ArgumentHelper arguments(args);

    bool printRowCol = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintRowCol(printRowCol);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::SetPrintRepeatRows) {
    NanScope();

    ArgumentHelper arguments(args);

    int rowFirst    = arguments.GetInt(0),
        rowLast     = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintRepeatRows(rowFirst, rowLast);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::SetPrintRepeatCols) {
    NanScope();

    ArgumentHelper arguments(args);

    int colFirst    = arguments.GetInt(0),
        colLast     = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintRepeatCols(colFirst, colLast);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::SetPrintArea) {
    NanScope();

    ArgumentHelper arguments(args);

    int rowFirst    = arguments.GetInt(0),
        rowLast     = arguments.GetInt(1),
        colFirst    = arguments.GetInt(2),
        colLast     = arguments.GetInt(3);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintArea(rowFirst, rowLast, colFirst, colLast);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::ClearPrintRepeats) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->clearPrintRepeats();

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::ClearPrintArea) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->clearPrintArea();

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::GetNamedRange) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value name(arguments.GetString(0));
    int scopeId = arguments.GetInt(1, libxl::SCOPE_UNDEFINED);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    int rowFirst, rowLast, colFirst, colLast;
    bool hidden;

    if (!that->GetWrapped()->getNamedRange(*name, &rowFirst, &rowLast, &colFirst, &colLast,
        scopeId, &hidden))
    {
        return util::ThrowLibxlError(that);
    }

    Local<Object> result = NanNew<Object>();
    result->Set(NanNew<String>("rowFirst"), NanNew<Integer>(rowFirst));
    result->Set(NanNew<String>("rowLast"),  NanNew<Integer>(rowLast));
    result->Set(NanNew<String>("colFirst"), NanNew<Integer>(colFirst));
    result->Set(NanNew<String>("colLast"),  NanNew<Integer>(colLast));
    result->Set(NanNew<String>("hidden"),   NanNew<Boolean>(hidden));

    NanReturnValue(result);
}


NAN_METHOD(Sheet::SetNamedRange) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value name(arguments.GetString(0));
    int rowFirst    = arguments.GetInt(1),
        rowLast     = arguments.GetInt(2),
        colFirst    = arguments.GetInt(3),
        colLast     = arguments.GetInt(4),
        scopeId     = arguments.GetInt(5, libxl::SCOPE_UNDEFINED);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setNamedRange(*name,
        rowFirst, rowLast, colFirst, colLast, static_cast<libxl::Scope>(scopeId)))
    {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::DelNamedRange) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value name(arguments.GetString(0));
    int scopeId = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->delNamedRange(*name,
        static_cast<libxl::Scope>(scopeId)))
    {
        util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::NamedRangeSize) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->namedRangeSize()));
}


NAN_METHOD(Sheet::NamedRange) {
    NanScope();

    ArgumentHelper arguments(args);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    int rowFirst, rowLast, colFirst, colLast, scopeId;
    bool hidden;
    const char* name = that->GetWrapped()->namedRange(
        index, &rowFirst, &rowLast, &colFirst, &colLast, &scopeId, &hidden);

    if (!name)
    {
        return util::ThrowLibxlError(that);
    }

    Local<Object> result = NanNew<Object>();
    result->Set(NanNew<String>("name"),     NanNew<String> (name));
    result->Set(NanNew<String>("rowFirst"), NanNew<Integer>(rowFirst));
    result->Set(NanNew<String>("rowLast"),  NanNew<Integer>(rowLast));
    result->Set(NanNew<String>("colFirst"), NanNew<Integer>(colFirst));
    result->Set(NanNew<String>("colLast"),  NanNew<Integer>(colLast));
    result->Set(NanNew<String>("scopeId"),  NanNew<Integer>(scopeId));
    result->Set(NanNew<String>("hidden"),   NanNew<Boolean>(hidden));

    NanReturnValue(result);
}


NAN_METHOD(Sheet::Name) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<String>(that->GetWrapped()->name()));
}


NAN_METHOD(Sheet::SetName) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value name(arguments.GetString(0));
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setName(*name);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::Protect) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->protect()));
}


NAN_METHOD(Sheet::SetProtect) {
    NanScope();

    ArgumentHelper arguments(args);

    bool protect = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setProtect(protect);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::RightToLeft) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->rightToLeft()));
}


NAN_METHOD(Sheet::SetRightToLeft) {
    NanScope();

    ArgumentHelper arguments(args);

    bool rightToLeft = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setRightToLeft(rightToLeft);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::Hidden) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->hidden()));
}


NAN_METHOD(Sheet::SetHidden) {
    NanScope();

    ArgumentHelper arguments(args);
    
    int state = arguments.GetInt(0, libxl::SHEETSTATE_HIDDEN);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setHidden(static_cast<libxl::SheetState>(state))) {
        util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::GetTopLeftView) {
    NanScope();

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    int row, col;
    that->GetWrapped()->getTopLeftView(&row, &col);

    Local<Object> result = NanNew<Object>();
    result->Set(NanNew<String>("row"), NanNew<Integer>(row));
    result->Set(NanNew<String>("col"), NanNew<Integer>(col));

    NanReturnValue(result);
}


NAN_METHOD(Sheet::SetTopLeftView) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0),
        col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setTopLeftView(row, col);

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::AddrToRowCol) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value addr(arguments.GetString(0));
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    int row = -1, col = -1;
    bool rowRelative = true, colRelative = true;

    that->GetWrapped()->addrToRowCol(*addr, &row, &col, &rowRelative, &colRelative);

    Local<Object> result = NanNew<Object>();
    result->Set(NanNew<String>("row"),          NanNew<Integer>(row));
    result->Set(NanNew<String>("col"),          NanNew<Integer>(col));
    result->Set(NanNew<String>("rowRelative"),  NanNew<Boolean>(rowRelative));
    result->Set(NanNew<String>("colRelative"),  NanNew<Boolean>(colRelative));

    NanReturnValue(result);
}


NAN_METHOD(Sheet::RowColToAddr) {
    NanScope();

    ArgumentHelper arguments(args);

    int row = arguments.GetInt(0),
        col = arguments.GetInt(1);
    bool rowRelative = arguments.GetBoolean(2, true),
         colRelative = arguments.GetBoolean(3, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    const char* addr = that->GetWrapped()->rowColToAddr(
        row, col, rowRelative, colRelative);

    if (!addr) {
        util::ThrowLibxlError(that);
    }

    NanReturnValue(NanNew<String>(addr));
}


// Init


void Sheet::Initialize(Handle<Object> exports) {
    using namespace libxl;

    NanScope();

    Local<FunctionTemplate> t = NanNew<FunctionTemplate>(util::StubConstructor);
    t->SetClassName(NanNew<String>("Sheet"));
    t->InstanceTemplate()->SetInternalFieldCount(1);

    BookWrapper::Initialize<Sheet>(t);

    NODE_SET_PROTOTYPE_METHOD(t, "cellType", CellType);
    NODE_SET_PROTOTYPE_METHOD(t, "isFormula", IsFormula);
    NODE_SET_PROTOTYPE_METHOD(t, "cellFormat", CellFormat);
    NODE_SET_PROTOTYPE_METHOD(t, "setCellFormat", SetCellFormat);
    NODE_SET_PROTOTYPE_METHOD(t, "readStr", ReadStr);
    NODE_SET_PROTOTYPE_METHOD(t, "readString", ReadStr);
    NODE_SET_PROTOTYPE_METHOD(t, "writeString", WriteStr);
    NODE_SET_PROTOTYPE_METHOD(t, "writeStr", WriteStr);
    NODE_SET_PROTOTYPE_METHOD(t, "readNum", ReadNum);
    NODE_SET_PROTOTYPE_METHOD(t, "writeNum", WriteNum);
    NODE_SET_PROTOTYPE_METHOD(t, "readBool", ReadBool);
    NODE_SET_PROTOTYPE_METHOD(t, "writeBool", WriteBool);
    NODE_SET_PROTOTYPE_METHOD(t, "readBlank", ReadBlank);
    NODE_SET_PROTOTYPE_METHOD(t, "writeBlank", WriteBlank);
    NODE_SET_PROTOTYPE_METHOD(t, "readFormula", ReadFormula);
    NODE_SET_PROTOTYPE_METHOD(t, "writeFormula", WriteFormula);
    NODE_SET_PROTOTYPE_METHOD(t, "readComment", ReadComment);
    NODE_SET_PROTOTYPE_METHOD(t, "writeComment", WriteComment);
    NODE_SET_PROTOTYPE_METHOD(t, "isDate", IsDate);
    NODE_SET_PROTOTYPE_METHOD(t, "readError", ReadError);
    NODE_SET_PROTOTYPE_METHOD(t, "colWidth", ColWidth);
    NODE_SET_PROTOTYPE_METHOD(t, "rowHeight", RowHeight);
    NODE_SET_PROTOTYPE_METHOD(t, "setCol", SetCol);
    NODE_SET_PROTOTYPE_METHOD(t, "setRow", SetRow);
    NODE_SET_PROTOTYPE_METHOD(t, "rowHidden", RowHidden);
    NODE_SET_PROTOTYPE_METHOD(t, "setRowHidden", SetRowHidden);
    NODE_SET_PROTOTYPE_METHOD(t, "colHidden", ColHidden);
    NODE_SET_PROTOTYPE_METHOD(t, "setColHidden", SetColHidden);
    NODE_SET_PROTOTYPE_METHOD(t, "getMerge", GetMerge);
    NODE_SET_PROTOTYPE_METHOD(t, "setMerge", SetMerge);
    NODE_SET_PROTOTYPE_METHOD(t, "delMerge", DelMerge);
    NODE_SET_PROTOTYPE_METHOD(t, "getHorPageBreak", GetHorPageBreak);
    NODE_SET_PROTOTYPE_METHOD(t, "getHorPageBreakSize", GetHorPageBreakSize);
    NODE_SET_PROTOTYPE_METHOD(t, "getVerPageBreak", GetVerPageBreak);
    NODE_SET_PROTOTYPE_METHOD(t, "getVerPageBreakSize", GetVerPageBreakSize);
    NODE_SET_PROTOTYPE_METHOD(t, "setHorPageBreak", SetHorPageBreak);
    NODE_SET_PROTOTYPE_METHOD(t, "setVerPageBreak", SetVerPageBreak);
    NODE_SET_PROTOTYPE_METHOD(t, "split", Split);
    NODE_SET_PROTOTYPE_METHOD(t, "groupRows", GroupRows);
    NODE_SET_PROTOTYPE_METHOD(t, "groupCols", GroupCols);
    NODE_SET_PROTOTYPE_METHOD(t, "groupSummaryBelow", GroupSummaryBelow);
    NODE_SET_PROTOTYPE_METHOD(t, "setGroupSummaryBelow", SetGroupSummaryBelow);
    NODE_SET_PROTOTYPE_METHOD(t, "groupSummaryRight", GroupSummaryRight);
    NODE_SET_PROTOTYPE_METHOD(t, "setGroupSummaryRight", SetGroupSummaryRight);
    NODE_SET_PROTOTYPE_METHOD(t, "insertRow", InsertRow);
    NODE_SET_PROTOTYPE_METHOD(t, "insertCol", InsertCol);
    NODE_SET_PROTOTYPE_METHOD(t, "removeRow", RemoveRow);
    NODE_SET_PROTOTYPE_METHOD(t, "removeCol", RemoveCol);
    NODE_SET_PROTOTYPE_METHOD(t, "copyCell", CopyCell);
    NODE_SET_PROTOTYPE_METHOD(t, "firstRow", FirstRow);
    NODE_SET_PROTOTYPE_METHOD(t, "lastRow", LastRow);
    NODE_SET_PROTOTYPE_METHOD(t, "firstCol", FirstCol);
    NODE_SET_PROTOTYPE_METHOD(t, "lastCol", LastCol);
    NODE_SET_PROTOTYPE_METHOD(t, "displayGridlines", DisplayGridlines);
    NODE_SET_PROTOTYPE_METHOD(t, "setDisplayGridlines", SetDisplayGridlines);
    NODE_SET_PROTOTYPE_METHOD(t, "printGridlines", PrintGridlines);
    NODE_SET_PROTOTYPE_METHOD(t, "setPrintGridlines", SetPrintGridlines);
    NODE_SET_PROTOTYPE_METHOD(t, "zoom", Zoom);
    NODE_SET_PROTOTYPE_METHOD(t, "setZoom", SetZoom);
    NODE_SET_PROTOTYPE_METHOD(t, "printZoom", PrintZoom);
    NODE_SET_PROTOTYPE_METHOD(t, "setPrintZoom", SetPrintZoom);
    NODE_SET_PROTOTYPE_METHOD(t, "getPrintFit", GetPrintFit);
    NODE_SET_PROTOTYPE_METHOD(t, "setPrintFit", SetPrintFit);
    NODE_SET_PROTOTYPE_METHOD(t, "landscape", Landscape);
    NODE_SET_PROTOTYPE_METHOD(t, "SetLandscape", SetLandscape);
    NODE_SET_PROTOTYPE_METHOD(t, "paper", Paper);
    NODE_SET_PROTOTYPE_METHOD(t, "setPaper", SetPaper);
    NODE_SET_PROTOTYPE_METHOD(t, "header", Header);
    NODE_SET_PROTOTYPE_METHOD(t, "setHeader", SetHeader);
    NODE_SET_PROTOTYPE_METHOD(t, "headerMargin", HeaderMargin);
    NODE_SET_PROTOTYPE_METHOD(t, "footer", Footer);
    NODE_SET_PROTOTYPE_METHOD(t, "setFooter", SetFooter);
    NODE_SET_PROTOTYPE_METHOD(t, "footerMargin", FooterMargin);
    NODE_SET_PROTOTYPE_METHOD(t, "hCenter", HCenter);
    NODE_SET_PROTOTYPE_METHOD(t, "setHCenter", SetHCenter);
    NODE_SET_PROTOTYPE_METHOD(t, "vCenter", VCenter);
    NODE_SET_PROTOTYPE_METHOD(t, "setVCenter", SetVCenter);
    NODE_SET_PROTOTYPE_METHOD(t, "marginLeft", MarginLeft);
    NODE_SET_PROTOTYPE_METHOD(t, "setMarginLeft", SetMarginLeft);
    NODE_SET_PROTOTYPE_METHOD(t, "marginRight", MarginRight);
    NODE_SET_PROTOTYPE_METHOD(t, "setMarginRight", SetMarginRight);
    NODE_SET_PROTOTYPE_METHOD(t, "marginTop", MarginTop);
    NODE_SET_PROTOTYPE_METHOD(t, "setMarginTop", SetMarginTop);
    NODE_SET_PROTOTYPE_METHOD(t, "marginBottom", MarginBottom);
    NODE_SET_PROTOTYPE_METHOD(t, "setMarginBottom", SetMarginBottom);
    NODE_SET_PROTOTYPE_METHOD(t, "printRowCol", PrintRowCol);
    NODE_SET_PROTOTYPE_METHOD(t, "setPrintRowCol", SetPrintRowCol);
    NODE_SET_PROTOTYPE_METHOD(t, "setPrintRepeatRows", SetPrintRepeatRows);
    NODE_SET_PROTOTYPE_METHOD(t, "setPrintRepeatCols", SetPrintRepeatCols);
    NODE_SET_PROTOTYPE_METHOD(t, "setPrintArea", SetPrintArea);
    NODE_SET_PROTOTYPE_METHOD(t, "clearPrintRepeats", ClearPrintRepeats);
    NODE_SET_PROTOTYPE_METHOD(t, "clearPrintArea", ClearPrintArea);
    NODE_SET_PROTOTYPE_METHOD(t, "getNamedRange", GetNamedRange);
    NODE_SET_PROTOTYPE_METHOD(t, "setNamedRange", SetNamedRange);
    NODE_SET_PROTOTYPE_METHOD(t, "delNamedRange", DelNamedRange);
    NODE_SET_PROTOTYPE_METHOD(t, "namedRangeSize", NamedRangeSize);
    NODE_SET_PROTOTYPE_METHOD(t, "namedRange", NamedRange);
    NODE_SET_PROTOTYPE_METHOD(t, "name", Name);
    NODE_SET_PROTOTYPE_METHOD(t, "setName",SetName);
    NODE_SET_PROTOTYPE_METHOD(t, "protect", Protect);
    NODE_SET_PROTOTYPE_METHOD(t, "setProtect", SetProtect);
    NODE_SET_PROTOTYPE_METHOD(t, "rightToLeft", RightToLeft);
    NODE_SET_PROTOTYPE_METHOD(t, "setRightToLeft", SetRightToLeft);
    NODE_SET_PROTOTYPE_METHOD(t, "hidden", Hidden);
    NODE_SET_PROTOTYPE_METHOD(t, "setHidden", SetHidden);
    NODE_SET_PROTOTYPE_METHOD(t, "getTopLeftView", GetTopLeftView);
    NODE_SET_PROTOTYPE_METHOD(t, "setTopLeftView", SetTopLeftView);
    NODE_SET_PROTOTYPE_METHOD(t, "addrToRowCol", AddrToRowCol);
    NODE_SET_PROTOTYPE_METHOD(t, "rowColToAddr", RowColToAddr);

    t->ReadOnlyPrototype();
    NanAssignPersistent(constructor, t->GetFunction());

    NODE_DEFINE_CONSTANT(exports, CELLTYPE_EMPTY);
    NODE_DEFINE_CONSTANT(exports, CELLTYPE_NUMBER);
    NODE_DEFINE_CONSTANT(exports, CELLTYPE_STRING);
    NODE_DEFINE_CONSTANT(exports, CELLTYPE_BOOLEAN);
    NODE_DEFINE_CONSTANT(exports, CELLTYPE_BLANK);
    NODE_DEFINE_CONSTANT(exports, CELLTYPE_ERROR);

    NODE_DEFINE_CONSTANT(exports, ERRORTYPE_NULL);
    NODE_DEFINE_CONSTANT(exports, ERRORTYPE_DIV_0);
    NODE_DEFINE_CONSTANT(exports, ERRORTYPE_VALUE);
    NODE_DEFINE_CONSTANT(exports, ERRORTYPE_REF);
    NODE_DEFINE_CONSTANT(exports, ERRORTYPE_NAME);
    NODE_DEFINE_CONSTANT(exports, ERRORTYPE_NUM);
    NODE_DEFINE_CONSTANT(exports, ERRORTYPE_NA);
    NODE_DEFINE_CONSTANT(exports, ERRORTYPE_NOERROR);

    NODE_DEFINE_CONSTANT(exports, PAPER_DEFAULT);
    NODE_DEFINE_CONSTANT(exports, PAPER_LETTER);
    NODE_DEFINE_CONSTANT(exports, PAPER_LETTERSMALL);
    NODE_DEFINE_CONSTANT(exports, PAPER_TABLOID);
    NODE_DEFINE_CONSTANT(exports, PAPER_LEDGER);
    NODE_DEFINE_CONSTANT(exports, PAPER_LEGAL);
    NODE_DEFINE_CONSTANT(exports, PAPER_STATEMENT);
    NODE_DEFINE_CONSTANT(exports, PAPER_EXECUTIVE);
    NODE_DEFINE_CONSTANT(exports, PAPER_A3);
    NODE_DEFINE_CONSTANT(exports, PAPER_A4);
    NODE_DEFINE_CONSTANT(exports, PAPER_A4SMALL);
    NODE_DEFINE_CONSTANT(exports, PAPER_A5);
    NODE_DEFINE_CONSTANT(exports, PAPER_B4);
    NODE_DEFINE_CONSTANT(exports, PAPER_B5);
    NODE_DEFINE_CONSTANT(exports, PAPER_FOLIO);
    NODE_DEFINE_CONSTANT(exports, PAPER_QUATRO);
    NODE_DEFINE_CONSTANT(exports, PAPER_10x14);
    NODE_DEFINE_CONSTANT(exports, PAPER_10x17);
    NODE_DEFINE_CONSTANT(exports, PAPER_NOTE);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_9);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_10);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_11);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_12);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_14);
    NODE_DEFINE_CONSTANT(exports, PAPER_C_SIZE);
    NODE_DEFINE_CONSTANT(exports, PAPER_D_SIZE);
    NODE_DEFINE_CONSTANT(exports, PAPER_E_SIZE);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_DL);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_C5);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_C3);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_C4);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_C6);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_C65);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_B4);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_B5);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_B6);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_MONARCH);
    NODE_DEFINE_CONSTANT(exports, PAPER_US_ENVELOPE);
    NODE_DEFINE_CONSTANT(exports, PAPER_FANFOLD);
    NODE_DEFINE_CONSTANT(exports, PAPER_GERMAN_STD_FANFOLD);
    NODE_DEFINE_CONSTANT(exports, PAPER_GERMAN_LEGAL_FANFOLD);
    NODE_DEFINE_CONSTANT(exports, PAPER_B4_ISO);
    NODE_DEFINE_CONSTANT(exports, PAPER_JAPANESE_POSTCARD);
    NODE_DEFINE_CONSTANT(exports, PAPER_9x11);
    NODE_DEFINE_CONSTANT(exports, PAPER_10x11);
    NODE_DEFINE_CONSTANT(exports, PAPER_15x11);
    NODE_DEFINE_CONSTANT(exports, PAPER_ENVELOPE_INVITE);
    NODE_DEFINE_CONSTANT(exports, PAPER_US_LETTER_EXTRA);
    NODE_DEFINE_CONSTANT(exports, PAPER_US_LEGAL_EXTRA);
    NODE_DEFINE_CONSTANT(exports, PAPER_US_TABLOID_EXTRA);
    NODE_DEFINE_CONSTANT(exports, PAPER_A4_EXTRA);
    NODE_DEFINE_CONSTANT(exports, PAPER_LETTER_TRANSVERSE);
    NODE_DEFINE_CONSTANT(exports, PAPER_A4_TRANSVERSE);
    NODE_DEFINE_CONSTANT(exports, PAPER_LETTER_EXTRA_TRANSVERSE);
    NODE_DEFINE_CONSTANT(exports, PAPER_SUPERA);
    NODE_DEFINE_CONSTANT(exports, PAPER_SUPERB);
    NODE_DEFINE_CONSTANT(exports, PAPER_US_LETTER_PLUS);
    NODE_DEFINE_CONSTANT(exports, PAPER_A4_PLUS);
    NODE_DEFINE_CONSTANT(exports, PAPER_A5_TRANSVERSE);
    NODE_DEFINE_CONSTANT(exports, PAPER_B5_TRANSVERSE);
    NODE_DEFINE_CONSTANT(exports, PAPER_A3_EXTRA);
    NODE_DEFINE_CONSTANT(exports, PAPER_A5_EXTRA);
    NODE_DEFINE_CONSTANT(exports, PAPER_B5_EXTRA);
    NODE_DEFINE_CONSTANT(exports, PAPER_A2);
    NODE_DEFINE_CONSTANT(exports, PAPER_A3_TRANSVERSE);
    NODE_DEFINE_CONSTANT(exports, PAPER_A3_EXTRA_TRANSVERSE);
    NODE_DEFINE_CONSTANT(exports, PAPER_JAPANESE_DOUBLE_POSTCARD);
    NODE_DEFINE_CONSTANT(exports, PAPER_A6);
    NODE_DEFINE_CONSTANT(exports, PAPER_JAPANESE_ENVELOPE_KAKU2);
    NODE_DEFINE_CONSTANT(exports, PAPER_JAPANESE_ENVELOPE_KAKU3);
    NODE_DEFINE_CONSTANT(exports, PAPER_JAPANESE_ENVELOPE_CHOU3);
    NODE_DEFINE_CONSTANT(exports, PAPER_JAPANESE_ENVELOPE_CHOU4);
    NODE_DEFINE_CONSTANT(exports, PAPER_LETTER_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_A3_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_A4_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_A5_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_B4_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_B5_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_JAPANESE_POSTCARD_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_DOUBLE_JAPANESE_POSTCARD_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_A6_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_JAPANESE_ENVELOPE_KAKU2_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_JAPANESE_ENVELOPE_KAKU3_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_JAPANESE_ENVELOPE_CHOU3_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_JAPANESE_ENVELOPE_CHOU4_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_B6);
    NODE_DEFINE_CONSTANT(exports, PAPER_B6_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_12x11);
    NODE_DEFINE_CONSTANT(exports, PAPER_JAPANESE_ENVELOPE_YOU4);
    NODE_DEFINE_CONSTANT(exports, PAPER_JAPANESE_ENVELOPE_YOU4_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC16K);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC32K);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC32K_BIG);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE1);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE2);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE3);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE4);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE5);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE6);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE7);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE8);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE9);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE10);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC16K_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC32K_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC32KBIG_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE1_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE2_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE3_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE4_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE5_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE6_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE7_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE8_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE9_ROTATED);
    NODE_DEFINE_CONSTANT(exports, PAPER_PRC_ENVELOPE10_ROTATED);

    NODE_DEFINE_CONSTANT(exports, SCOPE_UNDEFINED);
    NODE_DEFINE_CONSTANT(exports, SCOPE_WORKBOOK);

    NODE_DEFINE_CONSTANT(exports, SHEETSTATE_VISIBLE);
    NODE_DEFINE_CONSTANT(exports, SHEETSTATE_HIDDEN);
    NODE_DEFINE_CONSTANT(exports, SHEETSTATE_VERYHIDDEN);    
}


}
