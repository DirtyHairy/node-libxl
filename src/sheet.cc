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
#include "async_worker.h"

using namespace v8;

namespace node_libxl {


// Lifecycle

Sheet::Sheet(libxl::Sheet* sheet, Handle<Value> book) :
    Wrapper<libxl::Sheet>(sheet),
    BookWrapper(book)
{}


Local<Object> Sheet::NewInstance(
    libxl::Sheet* libxlSheet,
    Handle<Value> book)
{
    Nan::EscapableHandleScope scope;

    Sheet* sheet = new Sheet(libxlSheet, book);

    Local<Object> that = util::CallStubConstructor(
        Nan::New(constructor)).As<Object>();

    sheet->Wrap(that);

    return scope.Escape(that);
}


// Wrappers


NAN_METHOD(Sheet::CellType) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    libxl::CellType cellType = that->GetWrapped()->cellType(row, col);
    if (cellType == libxl::CELLTYPE_ERROR) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(Nan::New<Integer>(cellType));
}


NAN_METHOD(Sheet::IsFormula) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->isFormula(row, col)));
}


NAN_METHOD(Sheet::CellFormat) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    libxl::Format* libxlFormat = that->GetWrapped()->cellFormat(row, col);
    if (!libxlFormat) {
        return util::ThrowLibxlError(that);
    }

    
    info.GetReturnValue().Set(Format::NewInstance(libxlFormat, that->GetBookHandle()));
}


NAN_METHOD(Sheet::SetCellFormat) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Format* format = arguments.GetWrapped<Format>(2);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);
    ASSERT_SAME_BOOK(that, format);

    that->GetWrapped()->setCellFormat(row, col, format->GetWrapped());

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::ReadStr) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);
    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Handle<Value> formatRef = info[2];
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    libxl::Format* libxlFormat = NULL;
    const char* value = that->GetWrapped()->readStr(row, col, &libxlFormat);
    if (!value) {
        return util::ThrowLibxlError(that);
    }

    if (formatRef->IsObject() && libxlFormat) {
        formatRef.As<Object>()->Set(Nan::New<String>("format").ToLocalChecked(),
            Format::NewInstance(libxlFormat, that->GetBookHandle()));
    }

    info.GetReturnValue().Set(Nan::New<String>(value).ToLocalChecked());
}


NAN_METHOD(Sheet::WriteStr) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    String::Utf8Value value(arguments.GetString(2));
    Format* format = arguments.GetWrapped<Format>(3, NULL);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);
    if (format) {
        ASSERT_SAME_BOOK(that, format);
    }

    if (!that->GetWrapped()->
            writeStr(row, col, *value, format ? format->GetWrapped() : NULL))
    {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::ReadNum) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);
    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Handle<Value> formatRef = info[2];
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    libxl::Format* libxlFormat = NULL;
    double value = that->GetWrapped()->readNum(row, col, &libxlFormat);
    
    if (formatRef->IsObject() && libxlFormat) {
        formatRef.As<Object>()->Set(Nan::New<String>("format").ToLocalChecked(),
            Format::NewInstance(libxlFormat, that->GetBookHandle()));
    }

    info.GetReturnValue().Set(Nan::New<Number>(value));
}


NAN_METHOD(Sheet::WriteNum) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    double value = arguments.GetDouble(2);
    Format* format = arguments.GetWrapped<Format>(3, NULL);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);
    if (format) {
        ASSERT_SAME_BOOK(that, format);
    }

    if (!that->GetWrapped()->
            writeNum(row, col, value, format ? format->GetWrapped() : NULL))
    {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::ReadBool) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);
    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Handle<Value> formatRef = info[2];
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    libxl::Format* libxlFormat = NULL;
    bool value = that->GetWrapped()->readBool(row, col, &libxlFormat);
    
    if (formatRef->IsObject() && libxlFormat) {
        formatRef.As<Object>()->Set(Nan::New<String>("format").ToLocalChecked(),
            Format::NewInstance(libxlFormat, that->GetBookHandle()));
    }

    info.GetReturnValue().Set(Nan::New<Boolean>(value));
}


NAN_METHOD(Sheet::WriteBool) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    bool value = arguments.GetBoolean(2);
    Format* format = arguments.GetWrapped<Format>(3, NULL);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);
    if (format) {
        ASSERT_SAME_BOOK(that, format);
    }

    if (!that->GetWrapped()->
            writeBool(row, col, value, format ? format->GetWrapped() : NULL))
    {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::ReadBlank) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);
    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Handle<Value> formatRef = info[2];
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    libxl::Format* libxlFormat = NULL;
    if (!that->GetWrapped()->readBlank(row, col, &libxlFormat) || !libxlFormat) {
        return util::ThrowLibxlError(that);
    }

    Local<Value> formatHandle = Format::NewInstance(libxlFormat,
        that->GetBookHandle());
    
    if (formatRef->IsObject() && libxlFormat) {
        formatRef.As<Object>()->Set(Nan::New<String>("format").ToLocalChecked(), formatHandle);
    }

    info.GetReturnValue().Set(formatHandle);
}


NAN_METHOD(Sheet::WriteBlank) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Format* format = arguments.GetWrapped<Format>(2);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);
    ASSERT_SAME_BOOK(that, format);

    if (!that->GetWrapped()->
            writeBlank(row, col, format ? format->GetWrapped() : NULL))
    {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::ReadFormula) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);
    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    Handle<Value> formatRef = info[2];
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    libxl::Format* libxlFormat = NULL;
    const char* value = that->GetWrapped()->readFormula(row, col, &libxlFormat);
    if (!value) {
        return util::ThrowLibxlError(that);
    }

    if (formatRef->IsObject() && libxlFormat) {
        formatRef.As<Object>()->Set(Nan::New<String>("format").ToLocalChecked(),
            Format::NewInstance(libxlFormat, that->GetBookHandle()));
    }

    info.GetReturnValue().Set(Nan::New<String>(value).ToLocalChecked());
}


NAN_METHOD(Sheet::WriteFormula) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    String::Utf8Value value(arguments.GetString(2));
    Format* format = arguments.GetWrapped<Format>(3, NULL);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);
    if (format) {
        ASSERT_SAME_BOOK(that, format);
    }

    if (!that->GetWrapped()->
            writeFormula(row, col, *value, format? format->GetWrapped() : NULL))
    {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::ReadComment) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);
    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    const char* value = that->GetWrapped()->readComment(row, col);
    if (!value) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(Nan::New<String>(value).ToLocalChecked());
}


NAN_METHOD(Sheet::WriteComment) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    String::Utf8Value value(arguments.GetString(2));
    String::Utf8Value author(arguments.GetString(3, ""));
    int width = arguments.GetInt(4, 129);
    int height = arguments.GetInt(5, 75);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (info[3]->IsString()) {
        that->GetWrapped()->writeComment(row, col, *value, *author,
            width, height);
    } else {
        that->GetWrapped()->writeComment(row, col, *value, NULL,
            width, height);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::ReadError) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    int col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->readError(row, col)));
}


NAN_METHOD(Sheet::IsDate) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0),
        col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->isDate(row, col)));
}


NAN_METHOD(Sheet::ColWidth) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int col = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->colWidth(col)));
}


NAN_METHOD(Sheet::RowHeight) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->rowHeight(row)));
}


NAN_METHOD(Sheet::SetCol) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int first = arguments.GetInt(0);
    int last = arguments.GetInt(1);
    double width = arguments.GetDouble(2);
    Format* format = arguments.GetWrapped<Format>(3, NULL);
    bool hidden = arguments.GetBoolean(4, false);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);
    if (format) {
        ASSERT_SAME_BOOK(that, format);
    }

    if (!that->GetWrapped()->setCol(first, last, width,
            format ? format->GetWrapped() : NULL, hidden))
    {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::SetRow) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    double height = arguments.GetDouble(1);
    Format* format = arguments.GetWrapped<Format>(2, NULL);
    bool hidden = arguments.GetBoolean(3, false);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);
    if (format) {
        ASSERT_SAME_BOOK(that, format);
    }

    if (!that->GetWrapped()->setRow(row, height,
        format ? format->GetWrapped() : NULL, hidden))
    {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::RowHidden) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->rowHidden(row)));
}


NAN_METHOD(Sheet::SetRowHidden) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    bool hidden = arguments.GetBoolean(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setRowHidden(row, hidden)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::ColHidden) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int col = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->colHidden(col)));
}


NAN_METHOD(Sheet::SetColHidden) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int col = arguments.GetInt(0);
    bool hidden = arguments.GetBoolean(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setColHidden(col, hidden)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::GetMerge) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0),
        col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    int rowFirst, rowLast, colFirst, colLast;

    if (!that->GetWrapped()->getMerge(row, col, &rowFirst, &rowLast,
        &colFirst, &colLast))
    {
        return util::ThrowLibxlError(that);
    }

    Local<Object> result = Nan::New<Object>();
    result->Set(Nan::New<String>("rowFirst").ToLocalChecked(), Nan::New<Integer>(rowFirst));
    result->Set(Nan::New<String>("rowLast").ToLocalChecked(),  Nan::New<Integer>(rowLast));
    result->Set(Nan::New<String>("colFirst").ToLocalChecked(), Nan::New<Integer>(colFirst));
    result->Set(Nan::New<String>("colLast").ToLocalChecked(),  Nan::New<Integer>(colLast));

    info.GetReturnValue().Set(result);
}


NAN_METHOD(Sheet::SetMerge) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int rowFirst = arguments.GetInt(0);
    int rowLast = arguments.GetInt(1);
    int colFirst = arguments.GetInt(2);
    int colLast = arguments.GetInt(3);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setMerge(rowFirst, rowLast, colFirst, colLast)) {
        return util::ThrowLibxlError(that);
    }
    
    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::DelMerge) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0),
        col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->delMerge(row, col)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::PictureSize) {
    Nan::HandleScope scope;
    
    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->pictureSize()));
}


NAN_METHOD(Sheet::GetPicture) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int sheetIndex = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    int rowTop, colLeft, rowBottom, colRight, width, height,
        offset_x, offset_y;
    int bookIndex = that->GetWrapped()->getPicture(sheetIndex,
        &rowTop, &colLeft, &rowBottom, &colRight, &width, &height, &offset_x,
        &offset_y);

    if (bookIndex == -1) {
        return util::ThrowLibxlError(that);
    }

    Local<Object> result = Nan::New<Object>();
    result->Set(Nan::New<String>("bookIndex").ToLocalChecked(),    Nan::New<Integer>(bookIndex));
    result->Set(Nan::New<String>("rowTop").ToLocalChecked(),       Nan::New<Integer>(rowTop));
    result->Set(Nan::New<String>("colLeft").ToLocalChecked(),      Nan::New<Integer>(colLeft));
    result->Set(Nan::New<String>("rowBottom").ToLocalChecked(),    Nan::New<Integer>(rowBottom));
    result->Set(Nan::New<String>("colRight").ToLocalChecked(),     Nan::New<Integer>(colRight));
    result->Set(Nan::New<String>("width").ToLocalChecked(),        Nan::New<Integer>(width));
    result->Set(Nan::New<String>("height").ToLocalChecked(),       Nan::New<Integer>(height));
    result->Set(Nan::New<String>("offset_x").ToLocalChecked(),     Nan::New<Integer>(offset_x));
    result->Set(Nan::New<String>("offset_y").ToLocalChecked(),     Nan::New<Integer>(offset_y));

    info.GetReturnValue().Set(result);
}


NAN_METHOD(Sheet::SetPicture) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0),
        col = arguments.GetInt(1),
        id  = arguments.GetInt(2);
    double scale = arguments.GetDouble(3, 1.);
    int offset_x = arguments.GetInt(4, 0),
        offset_y = arguments.GetInt(5, 0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPicture(row, col, id, scale, offset_x, offset_y);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::SetPicture2) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row         = arguments.GetInt(0),
        col         = arguments.GetInt(1),
        id          = arguments.GetInt(2),
        width       = arguments.GetInt(3, -1),
        height      = arguments.GetInt(4, -1),
        offset_x    = arguments.GetInt(5, 0),
        offset_y    = arguments.GetInt(6, 0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPicture2(row, col, id, width, height, offset_x,
        offset_y);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::GetHorPageBreak) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->getHorPageBreak(index)));
}


NAN_METHOD(Sheet::GetHorPageBreakSize) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->getHorPageBreakSize()));
}


NAN_METHOD(Sheet::GetVerPageBreak) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);
    
    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->getVerPageBreak(index)));
}


NAN_METHOD(Sheet::GetVerPageBreakSize) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->getVerPageBreakSize()));
}


NAN_METHOD(Sheet::SetHorPageBreak) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0);
    bool pagebreak = arguments.GetBoolean(1, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setHorPageBreak(row, pagebreak)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::SetVerPageBreak) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int col = arguments.GetInt(0);
    bool pagebreak = arguments.GetBoolean(1, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setVerPageBreak(col, pagebreak)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::Split) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0),
        col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->split(row, col);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::GroupRows) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int rowFirst = arguments.GetInt(0),
        rowLast = arguments.GetInt(1);
    bool collapsed = arguments.GetBoolean(2, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->groupRows(rowFirst, rowLast, collapsed)) {
        return util::ThrowLibxlError(that);
    };

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::GroupCols) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int colFirst = arguments.GetInt(0),
        colLast = arguments.GetInt(1);
    bool collapsed = arguments.GetBoolean(2, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->groupCols(colFirst, colLast, collapsed)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::GroupSummaryBelow) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->groupSummaryBelow()));
}


NAN_METHOD(Sheet::SetGroupSummaryBelow) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool summaryBelow = arguments.GetBoolean(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setGroupSummaryBelow(summaryBelow);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::GroupSummaryRight) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->groupSummaryRight()));
}


NAN_METHOD(Sheet::SetGroupSummaryRight) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool summaryRight = arguments.GetBoolean(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setGroupSummaryRight(summaryRight);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::Clear) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int rowFirst    = arguments.GetInt(0, 0),
        rowLast     = arguments.GetInt(1, 65535),
        colFirst    = arguments.GetInt(2, 0),
        colLast     = arguments.GetInt(3, 255);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->clear(rowFirst, rowLast, colFirst, colLast);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::InsertRow) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int rowFirst    = arguments.GetInt(0),
        rowLast     = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->insertRow(rowFirst, rowLast)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::InsertRowAsync) {
    class Worker : public AsyncWorker<Sheet> {
        public:
            Worker(Nan::Callback* callback, Local<Object> that, int rowFirst,
                    int rowLast) :
                AsyncWorker<Sheet>(callback, that),
                rowFirst(rowFirst),
                rowLast(rowLast)
            {}

            virtual void Execute() {
                if (!that->GetWrapped()->insertRow(rowFirst, rowLast)) {
                    RaiseLibxlError();
                }
            }

        private:
            int rowFirst, rowLast;
    };

    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int rowFirst    = arguments.GetInt(0),
        rowLast     = arguments.GetInt(1);
    Handle<Function> callback = arguments.GetFunction(2);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    Nan::AsyncQueueWorker(new Worker(new Nan::Callback(callback),
        info.This(), rowFirst, rowLast));

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::InsertCol) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int colFirst    = arguments.GetInt(0),
        colLast     = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->insertCol(colFirst, colLast)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::InsertColAsync) {
    class Worker : public AsyncWorker<Sheet> {
        public:
            Worker(Nan::Callback* callback, Local<Object> that,
                    int colFirst, int colLast) :
                AsyncWorker<Sheet>(callback, that),
                colFirst(colFirst),
                colLast(colLast)
            {}

        virtual void Execute() {
            if (!that->GetWrapped()->insertCol(colFirst, colLast)) {
                RaiseLibxlError();
            }
        }

        private:
            int colFirst, colLast;
    };

    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int colFirst    = arguments.GetInt(0),
        colLast     = arguments.GetInt(1);
    Handle<Function> callback = arguments.GetFunction(2);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    Nan::AsyncQueueWorker(new Worker(new Nan::Callback(callback), info.This(),
        colFirst, colLast));

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::RemoveRow) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int rowFirst    = arguments.GetInt(0),
        rowLast     = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->removeRow(rowFirst, rowLast)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::RemoveRowAsync) {
    class Worker : public AsyncWorker<Sheet> {
        public:
            Worker(Nan::Callback* callback, Local<Object> that,
                    int rowFirst, int rowLast) :
                AsyncWorker<Sheet>(callback, that),
                rowFirst(rowFirst),
                rowLast(rowLast)
            {}

            virtual void Execute() {
                if (!that->GetWrapped()->removeRow(rowFirst, rowLast)) {
                    RaiseLibxlError();
                }
            }

        private:
            int rowFirst, rowLast;
    };

    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int rowFirst = arguments.GetInt(0),
        rowLast = arguments.GetInt(1);
    Handle<Function> callback = arguments.GetFunction(2);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    Nan::AsyncQueueWorker(new Worker(new Nan::Callback(callback), info.This(),
        rowFirst, rowLast));

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::RemoveColAsync) {
    class Worker : public AsyncWorker<Sheet> {
        public:
            Worker(Nan::Callback* callback, Local<Object> that,
                    int colFirst, int colLast) :
                AsyncWorker<Sheet>(callback, that),
                colFirst(colFirst),
                colLast(colLast)
            {}

            virtual void Execute() {
                if (!that->GetWrapped()->removeCol(colFirst, colLast)) {
                    RaiseLibxlError();
                }
            }

        private:
            int colFirst, colLast;
    };

    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int colFirst = arguments.GetInt(0),
        colLast = arguments.GetInt(1);
    Handle<Function> callback = arguments.GetFunction(2);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    Nan::AsyncQueueWorker(new Worker(new Nan::Callback(callback), info.This(),
        colFirst, colLast));

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::RemoveCol) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int colFirst    = arguments.GetInt(0),
        colLast     = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->removeCol(colFirst, colLast)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::CopyCell) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int rowSrc = arguments.GetInt(0),
        colSrc = arguments.GetInt(1),
        rowDst = arguments.GetInt(2),
        colDst = arguments.GetInt(3);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->copyCell(rowSrc, colSrc, rowDst, colDst)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::FirstRow) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->firstRow()));
}


NAN_METHOD(Sheet::LastRow) {
    Nan::HandleScope scope;


    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->lastRow()));
}


NAN_METHOD(Sheet::FirstCol) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->firstCol()));
}

NAN_METHOD(Sheet::LastCol) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->lastCol()));
}


NAN_METHOD(Sheet::DisplayGridlines) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->displayGridlines()));
}


NAN_METHOD(Sheet::SetDisplayGridlines) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool displayGridlines = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);
    
    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setDisplayGridlines(displayGridlines);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::PrintGridlines) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->printGridlines()));
}


NAN_METHOD(Sheet::SetPrintGridlines) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool printGridlines = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintGridlines(printGridlines);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::Zoom) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->zoom()));
}


NAN_METHOD(Sheet::SetZoom) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int zoom = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setZoom(zoom);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::PrintZoom) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->printZoom()));
}


NAN_METHOD(Sheet::SetPrintZoom) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int zoom = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintZoom(zoom);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::GetPrintFit) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    int wPages, hPages;

    if (that->GetWrapped()->getPrintFit(&wPages, &hPages)) {
        Local<Object> result = Nan::New<Object>();

        result->Set(Nan::New<String>("wPages").ToLocalChecked(),   Nan::New<Integer>(wPages));
        result->Set(Nan::New<String>("hPages").ToLocalChecked(),   Nan::New<Integer>(hPages));

        info.GetReturnValue().Set(result);
    } else {
        info.GetReturnValue().Set(Nan::False());
    }
}


NAN_METHOD(Sheet::SetPrintFit) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int wPages = arguments.GetInt(0, 1),
        hPages = arguments.GetInt(1, 1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintFit(wPages, hPages);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::Landscape) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->landscape()));
}


NAN_METHOD(Sheet::SetLandscape) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool landscape = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setLandscape(landscape);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::Paper) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->paper()));
}


NAN_METHOD(Sheet::SetPaper) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int paper = arguments.GetInt(0, libxl::PAPER_DEFAULT);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPaper(static_cast<libxl::Paper>(paper));

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::Header) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<String>(that->GetWrapped()->header()).ToLocalChecked());
}


NAN_METHOD(Sheet::SetHeader) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    String::Utf8Value header(arguments.GetString(0));
    double margin = arguments.GetDouble(1, 0.5);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setHeader(*header, margin)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::HeaderMargin) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->headerMargin()));
}



NAN_METHOD(Sheet::Footer) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<String>(that->GetWrapped()->footer()).ToLocalChecked());
}


NAN_METHOD(Sheet::SetFooter) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    String::Utf8Value footer(arguments.GetString(0));
    double margin = arguments.GetDouble(1, 0.5);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setFooter(*footer, margin)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::FooterMargin) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->footerMargin()));
}


NAN_METHOD(Sheet::HCenter) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->hCenter()));
}


NAN_METHOD(Sheet::SetHCenter) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool center = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setHCenter(center);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::VCenter) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->vCenter()));
}


NAN_METHOD(Sheet::SetVCenter) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool center = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setVCenter(center);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::MarginLeft) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->marginLeft()));
}


NAN_METHOD(Sheet::SetMarginLeft) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    double margin = arguments.GetDouble(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setMarginLeft(margin);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::MarginRight) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->marginRight()));
}


NAN_METHOD(Sheet::SetMarginRight) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    double margin = arguments.GetDouble(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setMarginRight(margin);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::MarginTop) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->marginTop()));
}


NAN_METHOD(Sheet::SetMarginTop) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    double margin = arguments.GetDouble(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setMarginTop(margin);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::MarginBottom) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->marginBottom()));
}


NAN_METHOD(Sheet::SetMarginBottom) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    double margin = arguments.GetDouble(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setMarginBottom(margin);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::PrintRowCol) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->printRowCol()));
}

NAN_METHOD(Sheet::SetPrintRowCol) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool printRowCol = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintRowCol(printRowCol);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::SetPrintRepeatRows) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int rowFirst    = arguments.GetInt(0),
        rowLast     = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintRepeatRows(rowFirst, rowLast);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::SetPrintRepeatCols) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int colFirst    = arguments.GetInt(0),
        colLast     = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintRepeatCols(colFirst, colLast);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::SetPrintArea) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int rowFirst    = arguments.GetInt(0),
        rowLast     = arguments.GetInt(1),
        colFirst    = arguments.GetInt(2),
        colLast     = arguments.GetInt(3);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPrintArea(rowFirst, rowLast, colFirst, colLast);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::ClearPrintRepeats) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->clearPrintRepeats();

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::ClearPrintArea) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->clearPrintArea();

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::GetNamedRange) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    String::Utf8Value name(arguments.GetString(0));
    int scopeId = arguments.GetInt(1, libxl::SCOPE_UNDEFINED);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    int rowFirst, rowLast, colFirst, colLast;
    bool hidden;

    if (!that->GetWrapped()->getNamedRange(*name, &rowFirst, &rowLast, &colFirst, &colLast,
        scopeId, &hidden))
    {
        return util::ThrowLibxlError(that);
    }

    Local<Object> result = Nan::New<Object>();
    result->Set(Nan::New<String>("rowFirst").ToLocalChecked(), Nan::New<Integer>(rowFirst));
    result->Set(Nan::New<String>("rowLast").ToLocalChecked(),  Nan::New<Integer>(rowLast));
    result->Set(Nan::New<String>("colFirst").ToLocalChecked(), Nan::New<Integer>(colFirst));
    result->Set(Nan::New<String>("colLast").ToLocalChecked(),  Nan::New<Integer>(colLast));
    result->Set(Nan::New<String>("hidden").ToLocalChecked(),   Nan::New<Boolean>(hidden));

    info.GetReturnValue().Set(result);
}


NAN_METHOD(Sheet::SetNamedRange) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    String::Utf8Value name(arguments.GetString(0));
    int rowFirst    = arguments.GetInt(1),
        rowLast     = arguments.GetInt(2),
        colFirst    = arguments.GetInt(3),
        colLast     = arguments.GetInt(4),
        scopeId     = arguments.GetInt(5, libxl::SCOPE_UNDEFINED);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setNamedRange(*name,
        rowFirst, rowLast, colFirst, colLast, static_cast<libxl::Scope>(scopeId)))
    {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::DelNamedRange) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    String::Utf8Value name(arguments.GetString(0));
    int scopeId = arguments.GetInt(1, libxl::SCOPE_UNDEFINED);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->delNamedRange(*name,
        static_cast<libxl::Scope>(scopeId)))
    {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::NamedRangeSize) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->namedRangeSize()));
}


NAN_METHOD(Sheet::NamedRange) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    int rowFirst, rowLast, colFirst, colLast, scopeId;
    bool hidden;
    const char* name = that->GetWrapped()->namedRange(
        index, &rowFirst, &rowLast, &colFirst, &colLast, &scopeId, &hidden);

    if (!name)
    {
        return util::ThrowLibxlError(that);
    }

    Local<Object> result = Nan::New<Object>();
    result->Set(Nan::New<String>("name").ToLocalChecked(),     Nan::New<String> (name).ToLocalChecked());
    result->Set(Nan::New<String>("rowFirst").ToLocalChecked(), Nan::New<Integer>(rowFirst));
    result->Set(Nan::New<String>("rowLast").ToLocalChecked(),  Nan::New<Integer>(rowLast));
    result->Set(Nan::New<String>("colFirst").ToLocalChecked(), Nan::New<Integer>(colFirst));
    result->Set(Nan::New<String>("colLast").ToLocalChecked(),  Nan::New<Integer>(colLast));
    result->Set(Nan::New<String>("scopeId").ToLocalChecked(),  Nan::New<Integer>(scopeId));
    result->Set(Nan::New<String>("hidden").ToLocalChecked(),   Nan::New<Boolean>(hidden));

    info.GetReturnValue().Set(result);
}


NAN_METHOD(Sheet::Name) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<String>(that->GetWrapped()->name()).ToLocalChecked());
}


NAN_METHOD(Sheet::SetName) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    String::Utf8Value name(arguments.GetString(0));
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setName(*name);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::Protect) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->protect()));
}


NAN_METHOD(Sheet::SetProtect) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool protect = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setProtect(protect);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::RightToLeft) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->rightToLeft()));
}


NAN_METHOD(Sheet::SetRightToLeft) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool rightToLeft = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setRightToLeft(rightToLeft);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::Hidden) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->hidden()));
}


NAN_METHOD(Sheet::SetHidden) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);
    
    int state = arguments.GetInt(0, libxl::SHEETSTATE_HIDDEN);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setHidden(static_cast<libxl::SheetState>(state))) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::GetTopLeftView) {
    Nan::HandleScope scope;

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    int row, col;
    that->GetWrapped()->getTopLeftView(&row, &col);

    Local<Object> result = Nan::New<Object>();
    result->Set(Nan::New<String>("row").ToLocalChecked(), Nan::New<Integer>(row));
    result->Set(Nan::New<String>("col").ToLocalChecked(), Nan::New<Integer>(col));

    info.GetReturnValue().Set(result);
}


NAN_METHOD(Sheet::SetTopLeftView) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0),
        col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setTopLeftView(row, col);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Sheet::AddrToRowCol) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    String::Utf8Value addr(arguments.GetString(0));
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    int row = -1, col = -1;
    bool rowRelative = true, colRelative = true;

    that->GetWrapped()->addrToRowCol(*addr, &row, &col, &rowRelative, &colRelative);

    Local<Object> result = Nan::New<Object>();
    result->Set(Nan::New<String>("row").ToLocalChecked(),          Nan::New<Integer>(row));
    result->Set(Nan::New<String>("col").ToLocalChecked(),          Nan::New<Integer>(col));
    result->Set(Nan::New<String>("rowRelative").ToLocalChecked(),  Nan::New<Boolean>(rowRelative));
    result->Set(Nan::New<String>("colRelative").ToLocalChecked(),  Nan::New<Boolean>(colRelative));

    info.GetReturnValue().Set(result);
}


NAN_METHOD(Sheet::RowColToAddr) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int row = arguments.GetInt(0),
        col = arguments.GetInt(1);
    bool rowRelative = arguments.GetBoolean(2, true),
         colRelative = arguments.GetBoolean(3, true);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(info.This());
    ASSERT_THIS(that);

    const char* addr = that->GetWrapped()->rowColToAddr(
        row, col, rowRelative, colRelative);

    if (!addr) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(Nan::New<String>(addr).ToLocalChecked());
}


// Init


void Sheet::Initialize(Handle<Object> exports) {
    using namespace libxl;

    Nan::HandleScope scope;

    Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(util::StubConstructor);
    t->SetClassName(Nan::New<String>("Sheet").ToLocalChecked());
    t->InstanceTemplate()->SetInternalFieldCount(1);

    BookWrapper::Initialize<Sheet>(t);

    Nan::SetPrototypeMethod(t, "cellType", CellType);
    Nan::SetPrototypeMethod(t, "isFormula", IsFormula);
    Nan::SetPrototypeMethod(t, "cellFormat", CellFormat);
    Nan::SetPrototypeMethod(t, "setCellFormat", SetCellFormat);
    Nan::SetPrototypeMethod(t, "readStr", ReadStr);
    Nan::SetPrototypeMethod(t, "readString", ReadStr);
    Nan::SetPrototypeMethod(t, "writeString", WriteStr);
    Nan::SetPrototypeMethod(t, "writeStr", WriteStr);
    Nan::SetPrototypeMethod(t, "readNum", ReadNum);
    Nan::SetPrototypeMethod(t, "writeNum", WriteNum);
    Nan::SetPrototypeMethod(t, "readBool", ReadBool);
    Nan::SetPrototypeMethod(t, "writeBool", WriteBool);
    Nan::SetPrototypeMethod(t, "readBlank", ReadBlank);
    Nan::SetPrototypeMethod(t, "writeBlank", WriteBlank);
    Nan::SetPrototypeMethod(t, "readFormula", ReadFormula);
    Nan::SetPrototypeMethod(t, "writeFormula", WriteFormula);
    Nan::SetPrototypeMethod(t, "readComment", ReadComment);
    Nan::SetPrototypeMethod(t, "writeComment", WriteComment);
    Nan::SetPrototypeMethod(t, "isDate", IsDate);
    Nan::SetPrototypeMethod(t, "readError", ReadError);
    Nan::SetPrototypeMethod(t, "colWidth", ColWidth);
    Nan::SetPrototypeMethod(t, "rowHeight", RowHeight);
    Nan::SetPrototypeMethod(t, "setCol", SetCol);
    Nan::SetPrototypeMethod(t, "setRow", SetRow);
    Nan::SetPrototypeMethod(t, "rowHidden", RowHidden);
    Nan::SetPrototypeMethod(t, "setRowHidden", SetRowHidden);
    Nan::SetPrototypeMethod(t, "colHidden", ColHidden);
    Nan::SetPrototypeMethod(t, "setColHidden", SetColHidden);
    Nan::SetPrototypeMethod(t, "getMerge", GetMerge);
    Nan::SetPrototypeMethod(t, "setMerge", SetMerge);
    Nan::SetPrototypeMethod(t, "delMerge", DelMerge);
    Nan::SetPrototypeMethod(t, "pictureSize", PictureSize);
    Nan::SetPrototypeMethod(t, "getPicture", GetPicture);
    Nan::SetPrototypeMethod(t, "setPicture", SetPicture);
    Nan::SetPrototypeMethod(t, "setPicture2", SetPicture2);
    Nan::SetPrototypeMethod(t, "getHorPageBreak", GetHorPageBreak);
    Nan::SetPrototypeMethod(t, "getHorPageBreakSize", GetHorPageBreakSize);
    Nan::SetPrototypeMethod(t, "getVerPageBreak", GetVerPageBreak);
    Nan::SetPrototypeMethod(t, "getVerPageBreakSize", GetVerPageBreakSize);
    Nan::SetPrototypeMethod(t, "setHorPageBreak", SetHorPageBreak);
    Nan::SetPrototypeMethod(t, "setVerPageBreak", SetVerPageBreak);
    Nan::SetPrototypeMethod(t, "split", Split);
    Nan::SetPrototypeMethod(t, "groupRows", GroupRows);
    Nan::SetPrototypeMethod(t, "groupCols", GroupCols);
    Nan::SetPrototypeMethod(t, "groupSummaryBelow", GroupSummaryBelow);
    Nan::SetPrototypeMethod(t, "setGroupSummaryBelow", SetGroupSummaryBelow);
    Nan::SetPrototypeMethod(t, "groupSummaryRight", GroupSummaryRight);
    Nan::SetPrototypeMethod(t, "setGroupSummaryRight", SetGroupSummaryRight);
    Nan::SetPrototypeMethod(t, "clear", Clear);
    Nan::SetPrototypeMethod(t, "insertRow", InsertRow);
    Nan::SetPrototypeMethod(t, "insertRowAsync", InsertRowAsync);
    Nan::SetPrototypeMethod(t, "insertCol", InsertCol);
    Nan::SetPrototypeMethod(t, "insertColAsync", InsertColAsync);
    Nan::SetPrototypeMethod(t, "removeRow", RemoveRow);
    Nan::SetPrototypeMethod(t, "removeRowAsync", RemoveRowAsync);
    Nan::SetPrototypeMethod(t, "removeCol", RemoveCol);
    Nan::SetPrototypeMethod(t, "removeColAsync", RemoveColAsync);
    Nan::SetPrototypeMethod(t, "copyCell", CopyCell);
    Nan::SetPrototypeMethod(t, "firstRow", FirstRow);
    Nan::SetPrototypeMethod(t, "lastRow", LastRow);
    Nan::SetPrototypeMethod(t, "firstCol", FirstCol);
    Nan::SetPrototypeMethod(t, "lastCol", LastCol);
    Nan::SetPrototypeMethod(t, "displayGridlines", DisplayGridlines);
    Nan::SetPrototypeMethod(t, "setDisplayGridlines", SetDisplayGridlines);
    Nan::SetPrototypeMethod(t, "printGridlines", PrintGridlines);
    Nan::SetPrototypeMethod(t, "setPrintGridlines", SetPrintGridlines);
    Nan::SetPrototypeMethod(t, "zoom", Zoom);
    Nan::SetPrototypeMethod(t, "setZoom", SetZoom);
    Nan::SetPrototypeMethod(t, "printZoom", PrintZoom);
    Nan::SetPrototypeMethod(t, "setPrintZoom", SetPrintZoom);
    Nan::SetPrototypeMethod(t, "getPrintFit", GetPrintFit);
    Nan::SetPrototypeMethod(t, "setPrintFit", SetPrintFit);
    Nan::SetPrototypeMethod(t, "landscape", Landscape);
    Nan::SetPrototypeMethod(t, "setLandscape", SetLandscape);
    Nan::SetPrototypeMethod(t, "paper", Paper);
    Nan::SetPrototypeMethod(t, "setPaper", SetPaper);
    Nan::SetPrototypeMethod(t, "header", Header);
    Nan::SetPrototypeMethod(t, "setHeader", SetHeader);
    Nan::SetPrototypeMethod(t, "headerMargin", HeaderMargin);
    Nan::SetPrototypeMethod(t, "footer", Footer);
    Nan::SetPrototypeMethod(t, "setFooter", SetFooter);
    Nan::SetPrototypeMethod(t, "footerMargin", FooterMargin);
    Nan::SetPrototypeMethod(t, "hCenter", HCenter);
    Nan::SetPrototypeMethod(t, "setHCenter", SetHCenter);
    Nan::SetPrototypeMethod(t, "vCenter", VCenter);
    Nan::SetPrototypeMethod(t, "setVCenter", SetVCenter);
    Nan::SetPrototypeMethod(t, "marginLeft", MarginLeft);
    Nan::SetPrototypeMethod(t, "setMarginLeft", SetMarginLeft);
    Nan::SetPrototypeMethod(t, "marginRight", MarginRight);
    Nan::SetPrototypeMethod(t, "setMarginRight", SetMarginRight);
    Nan::SetPrototypeMethod(t, "marginTop", MarginTop);
    Nan::SetPrototypeMethod(t, "setMarginTop", SetMarginTop);
    Nan::SetPrototypeMethod(t, "marginBottom", MarginBottom);
    Nan::SetPrototypeMethod(t, "setMarginBottom", SetMarginBottom);
    Nan::SetPrototypeMethod(t, "printRowCol", PrintRowCol);
    Nan::SetPrototypeMethod(t, "setPrintRowCol", SetPrintRowCol);
    Nan::SetPrototypeMethod(t, "setPrintRepeatRows", SetPrintRepeatRows);
    Nan::SetPrototypeMethod(t, "setPrintRepeatCols", SetPrintRepeatCols);
    Nan::SetPrototypeMethod(t, "setPrintArea", SetPrintArea);
    Nan::SetPrototypeMethod(t, "clearPrintRepeats", ClearPrintRepeats);
    Nan::SetPrototypeMethod(t, "clearPrintArea", ClearPrintArea);
    Nan::SetPrototypeMethod(t, "getNamedRange", GetNamedRange);
    Nan::SetPrototypeMethod(t, "setNamedRange", SetNamedRange);
    Nan::SetPrototypeMethod(t, "delNamedRange", DelNamedRange);
    Nan::SetPrototypeMethod(t, "namedRangeSize", NamedRangeSize);
    Nan::SetPrototypeMethod(t, "namedRange", NamedRange);
    Nan::SetPrototypeMethod(t, "name", Name);
    Nan::SetPrototypeMethod(t, "setName",SetName);
    Nan::SetPrototypeMethod(t, "protect", Protect);
    Nan::SetPrototypeMethod(t, "setProtect", SetProtect);
    Nan::SetPrototypeMethod(t, "rightToLeft", RightToLeft);
    Nan::SetPrototypeMethod(t, "setRightToLeft", SetRightToLeft);
    Nan::SetPrototypeMethod(t, "hidden", Hidden);
    Nan::SetPrototypeMethod(t, "setHidden", SetHidden);
    Nan::SetPrototypeMethod(t, "getTopLeftView", GetTopLeftView);
    Nan::SetPrototypeMethod(t, "setTopLeftView", SetTopLeftView);
    Nan::SetPrototypeMethod(t, "addrToRowCol", AddrToRowCol);
    Nan::SetPrototypeMethod(t, "rowColToAddr", RowColToAddr);

    t->ReadOnlyPrototype();
    constructor.Reset(t->GetFunction());

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
