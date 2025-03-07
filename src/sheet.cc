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

#include "argument_helper.h"
#include "assert.h"
#include "async_worker.h"
#include "auto_filter.h"
#include "conditional_formatting.h"
#include "form_control.h"
#include "format.h"
#include "rich_string.h"
#include "util.h"

using namespace v8;

#define ASSERT_SHEET(sheet)                                                     \
    ASSERT_THIS(sheet);                                                         \
    if (!::node_libxl::util::GetBook(sheet)->IsValidSheet(sheet->wrappedSheet)) \
        return (Nan::ThrowError("sheet has been discarded and is no longer valid"));

namespace node_libxl {

    // Lifecycle

    Sheet::Sheet(libxl::Sheet* sheet, Local<Value> book)
        : Wrapper<libxl::Sheet, Sheet>(sheet), BookHolder(book), wrappedSheet(sheet) {}

    Local<Object> Sheet::NewInstance(libxl::Sheet* libxlSheet, Local<Value> book) {
        Nan::EscapableHandleScope scope;

        Sheet* sheet = new Sheet(libxlSheet, book);

        Local<Object> that = util::CallStubConstructor(Nan::New(constructor)).As<Object>();

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

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

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

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->isFormula(row, col)));
    }

    NAN_METHOD(Sheet::CellFormat) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

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

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        ASSERT_SAME_BOOK(that, format);

        that->GetWrapped()->setCellFormat(row, col, format->GetWrapped());

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::ReadStr) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        Local<Value> formatRef = info[2];
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        libxl::Format* libxlFormat = NULL;
        const char* value = that->GetWrapped()->readStr(row, col, &libxlFormat);
        if (!value) {
            return util::ThrowLibxlError(that);
        }

        if (formatRef->IsObject() && libxlFormat) {
            Nan::Set(formatRef.As<Object>(), Nan::New<String>("format").ToLocalChecked(),
                     Format::NewInstance(libxlFormat, that->GetBookHandle()));
        }

        info.GetReturnValue().Set(Nan::New<String>(value).ToLocalChecked());
    }

    NAN_METHOD(Sheet::WriteStr) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        CSNanUtf8Value(value, arguments.GetString(2));
        Format* format = arguments.GetWrapped<Format>(3, NULL);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        if (format) {
            ASSERT_SAME_BOOK(that, format);
        }

        if (!that->GetWrapped()->writeStr(row, col, *value, format ? format->GetWrapped() : NULL)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::ReadRichStr) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        Local<Value> formatRef = info[2];

        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        libxl::Format* libxlFormat = nullptr;
        libxl::RichString* richString = that->GetWrapped()->readRichStr(row, col, &libxlFormat);
        if (!richString) {
            return util::ThrowLibxlError(that);
        }

        if (formatRef->IsObject() && libxlFormat) {
            Nan::Set(formatRef.As<Object>(), Nan::New<String>("format").ToLocalChecked(),
                     Format::NewInstance(libxlFormat, that->GetBookHandle()));
        }

        info.GetReturnValue().Set(RichString::NewInstance(richString, that->GetBookHandle()));
    }

    NAN_METHOD(Sheet::WriteRichStr) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        RichString* richString = arguments.GetWrapped<RichString>(2);
        Format* format = arguments.GetWrapped<Format>(3, nullptr);

        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        ASSERT_SAME_BOOK(that, richString);
        if (format) ASSERT_SAME_BOOK(that, format);

        if (!that->GetWrapped()->writeRichStr(row, col, richString->GetWrapped(),
                                              format ? format->GetWrapped() : nullptr)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::ReadNum) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        Local<Value> formatRef = info[2];
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        libxl::Format* libxlFormat = NULL;
        double value = that->GetWrapped()->readNum(row, col, &libxlFormat);

        if (formatRef->IsObject() && libxlFormat) {
            Nan::Set(formatRef.As<Object>(), Nan::New<String>("format").ToLocalChecked(),
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

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        if (format) {
            ASSERT_SAME_BOOK(that, format);
        }

        if (!that->GetWrapped()->writeNum(row, col, value, format ? format->GetWrapped() : NULL)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::ReadBool) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        Local<Value> formatRef = info[2];
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        libxl::Format* libxlFormat = NULL;
        bool value = that->GetWrapped()->readBool(row, col, &libxlFormat);

        if (formatRef->IsObject() && libxlFormat) {
            Nan::Set(formatRef.As<Object>(), Nan::New<String>("format").ToLocalChecked(),
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

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        if (format) {
            ASSERT_SAME_BOOK(that, format);
        }

        if (!that->GetWrapped()->writeBool(row, col, value, format ? format->GetWrapped() : NULL)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::ReadBlank) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        Local<Value> formatRef = info[2];
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        libxl::Format* libxlFormat = NULL;
        if (!that->GetWrapped()->readBlank(row, col, &libxlFormat) || !libxlFormat) {
            return util::ThrowLibxlError(that);
        }

        Local<Value> formatHandle = Format::NewInstance(libxlFormat, that->GetBookHandle());

        if (formatRef->IsObject() && libxlFormat) {
            Nan::Set(formatRef.As<Object>(), Nan::New<String>("format").ToLocalChecked(),
                     formatHandle);
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

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        ASSERT_SAME_BOOK(that, format);

        if (!that->GetWrapped()->writeBlank(row, col, format ? format->GetWrapped() : NULL)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::ReadFormula) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        Local<Value> formatRef = info[2];
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        libxl::Format* libxlFormat = NULL;
        const char* value = that->GetWrapped()->readFormula(row, col, &libxlFormat);
        if (!value) {
            return util::ThrowLibxlError(that);
        }

        if (formatRef->IsObject() && libxlFormat) {
            Nan::Set(formatRef.As<Object>(), Nan::New<String>("format").ToLocalChecked(),
                     Format::NewInstance(libxlFormat, that->GetBookHandle()));
        }

        info.GetReturnValue().Set(Nan::New<String>(value).ToLocalChecked());
    }

    NAN_METHOD(Sheet::WriteFormula) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        CSNanUtf8Value(value, arguments.GetString(2));
        Format* format = arguments.GetWrapped<Format>(3, NULL);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        if (format) {
            ASSERT_SAME_BOOK(that, format);
        }

        if (!that->GetWrapped()->writeFormula(row, col, *value,
                                              format ? format->GetWrapped() : NULL)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::WriteFormulaNum) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        CSNanUtf8Value(expr, arguments.GetString(2));
        double value = arguments.GetDouble(3);
        Format* format = arguments.GetWrapped<Format>(4, NULL);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        if (format) {
            ASSERT_SAME_BOOK(that, format);
        }

        if (!that->GetWrapped()->writeFormulaNum(row, col, *expr, value,
                                                 format ? format->GetWrapped() : NULL)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::WriteFormulaStr) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        CSNanUtf8Value(expr, arguments.GetString(2));
        CSNanUtf8Value(value, arguments.GetString(3));
        Format* format = arguments.GetWrapped<Format>(4, NULL);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        if (format) {
            ASSERT_SAME_BOOK(that, format);
        }

        if (!that->GetWrapped()->writeFormulaStr(row, col, *expr, *value,
                                                 format ? format->GetWrapped() : NULL)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::WriteFormulaBool) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        CSNanUtf8Value(expr, arguments.GetString(2));
        bool value = arguments.GetBoolean(3);
        Format* format = arguments.GetWrapped<Format>(4, NULL);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        if (format) {
            ASSERT_SAME_BOOK(that, format);
        }

        if (!that->GetWrapped()->writeFormulaBool(row, col, *expr, value,
                                                  format ? format->GetWrapped() : NULL)) {
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

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

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
        CSNanUtf8Value(value, arguments.GetString(2));
        CSNanUtf8Value(author, arguments.GetString(3, ""));
        int width = arguments.GetInt(4, 129);
        int height = arguments.GetInt(5, 75);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (info[3]->IsString()) {
            that->GetWrapped()->writeComment(row, col, *value, *author, width, height);
        } else {
            that->GetWrapped()->writeComment(row, col, *value, NULL, width, height);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::RemoveComment) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->removeComment(row, col);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::ReadError) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->readError(row, col)));
    }

    NAN_METHOD(Sheet::WriteError) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        int error = arguments.GetInt(2);
        Format* format = arguments.GetWrapped<Format>(3, nullptr);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        if (format) {
            ASSERT_SAME_BOOK(that, format);
        }

        that->GetWrapped()->writeError(row, col, static_cast<libxl::ErrorType>(error),
                                       format ? format->GetWrapped() : nullptr);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::IsDate) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0), col = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->isDate(row, col)));
    }

    NAN_METHOD(Sheet::IsRichStr) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0), col = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->isRichStr(row, col)));
    }

    NAN_METHOD(Sheet::ColWidth) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int col = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->colWidth(col)));
    }

    NAN_METHOD(Sheet::RowHeight) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->rowHeight(row)));
    }

    NAN_METHOD(Sheet::ColWidthPx) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int col = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->colWidthPx(col)));
    }

    NAN_METHOD(Sheet::RowHeightPx) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->rowHeightPx(row)));
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

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        if (format) {
            ASSERT_SAME_BOOK(that, format);
        }

        if (!that->GetWrapped()->setCol(first, last, width, format ? format->GetWrapped() : NULL,
                                        hidden)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::SetColPx) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int first = arguments.GetInt(0);
        int last = arguments.GetInt(1);
        int width = arguments.GetInt(2);
        Format* format = arguments.GetWrapped<Format>(3, NULL);
        bool hidden = arguments.GetBoolean(4, false);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        if (format) {
            ASSERT_SAME_BOOK(that, format);
        }

        if (!that->GetWrapped()->setColPx(first, last, width, format ? format->GetWrapped() : NULL,
                                          hidden)) {
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

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        if (format) {
            ASSERT_SAME_BOOK(that, format);
        }

        if (!that->GetWrapped()->setRow(row, height, format ? format->GetWrapped() : NULL,
                                        hidden)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::SetRowPx) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        int height = arguments.GetInt(1);
        Format* format = arguments.GetWrapped<Format>(2, NULL);
        bool hidden = arguments.GetBoolean(3, false);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);
        if (format) {
            ASSERT_SAME_BOOK(that, format);
        }

        if (!that->GetWrapped()->setRowPx(row, height, format ? format->GetWrapped() : NULL,
                                          hidden)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::ColFormat) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int col = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        libxl::Format* format = that->GetWrapped()->colFormat(col);
        if (!format) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Format::NewInstance(format, that->GetBookHandle()));
    }

    NAN_METHOD(Sheet::RowFormat) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        libxl::Format* format = that->GetWrapped()->rowFormat(row);
        if (!format) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Format::NewInstance(format, that->GetBookHandle()));
    }

    NAN_METHOD(Sheet::RowHidden) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->rowHidden(row)));
    }

    NAN_METHOD(Sheet::SetRowHidden) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        bool hidden = arguments.GetBoolean(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

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

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->colHidden(col)));
    }

    NAN_METHOD(Sheet::SetColHidden) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int col = arguments.GetInt(0);
        bool hidden = arguments.GetBoolean(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->setColHidden(col, hidden)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::DefaultRowHeight) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->defaultRowHeight()));
    }

    NAN_METHOD(Sheet::SetDefaultRowHeight) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        double height = arguments.GetDouble(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setDefaultRowHeight(height);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::GetMerge) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0), col = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int rowFirst, rowLast, colFirst, colLast;

        if (!that->GetWrapped()->getMerge(row, col, &rowFirst, &rowLast, &colFirst, &colLast)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("rowFirst").ToLocalChecked(),
                 Nan::New<Integer>(rowFirst));
        Nan::Set(result, Nan::New<String>("rowLast").ToLocalChecked(), Nan::New<Integer>(rowLast));
        Nan::Set(result, Nan::New<String>("colFirst").ToLocalChecked(),
                 Nan::New<Integer>(colFirst));
        Nan::Set(result, Nan::New<String>("colLast").ToLocalChecked(), Nan::New<Integer>(colLast));

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

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->setMerge(rowFirst, rowLast, colFirst, colLast)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::DelMerge) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0), col = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->delMerge(row, col)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::MergeSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->mergeSize()));
    }

    NAN_METHOD(Sheet::Merge) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int index = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int rowFirst, rowLast, colFirst, colLast;

        if (!that->GetWrapped()->merge(index, &rowFirst, &rowLast, &colFirst, &colLast)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("rowFirst").ToLocalChecked(),
                 Nan::New<Integer>(rowFirst));
        Nan::Set(result, Nan::New<String>("rowLast").ToLocalChecked(), Nan::New<Integer>(rowLast));
        Nan::Set(result, Nan::New<String>("colFirst").ToLocalChecked(),
                 Nan::New<Integer>(colFirst));
        Nan::Set(result, Nan::New<String>("colLast").ToLocalChecked(), Nan::New<Integer>(colLast));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::DelMergeByIndex) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int index = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->delMergeByIndex(index)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::PictureSize) {
        Nan::HandleScope scope;

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->pictureSize()));
    }

    NAN_METHOD(Sheet::GetPicture) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int sheetIndex = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int rowTop, colLeft, rowBottom, colRight, width, height, offset_x, offset_y;
        const char* linkPath;
        int bookIndex =
            that->GetWrapped()->getPicture(sheetIndex, &rowTop, &colLeft, &rowBottom, &colRight,
                                           &width, &height, &offset_x, &offset_y, &linkPath);

        if (bookIndex == -1) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("bookIndex").ToLocalChecked(),
                 Nan::New<Integer>(bookIndex));
        Nan::Set(result, Nan::New<String>("rowTop").ToLocalChecked(), Nan::New<Integer>(rowTop));
        Nan::Set(result, Nan::New<String>("colLeft").ToLocalChecked(), Nan::New<Integer>(colLeft));
        Nan::Set(result, Nan::New<String>("rowBottom").ToLocalChecked(),
                 Nan::New<Integer>(rowBottom));
        Nan::Set(result, Nan::New<String>("colRight").ToLocalChecked(),
                 Nan::New<Integer>(colRight));
        Nan::Set(result, Nan::New<String>("width").ToLocalChecked(), Nan::New<Integer>(width));
        Nan::Set(result, Nan::New<String>("height").ToLocalChecked(), Nan::New<Integer>(height));
        Nan::Set(result, Nan::New<String>("offset_x").ToLocalChecked(),
                 Nan::New<Integer>(offset_x));
        Nan::Set(result, Nan::New<String>("offset_y").ToLocalChecked(),
                 Nan::New<Integer>(offset_y));
        if (linkPath) {
            Nan::Set(result, Nan::New<String>("linkPath").ToLocalChecked(),
                     Nan::New<String>(linkPath).ToLocalChecked());
        }

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::RemovePicture) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0), col = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->removePicture(row, col)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::RemovePictureByIndex) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int index = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->removePictureByIndex(index)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::SetPicture) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0), col = arguments.GetInt(1), id = arguments.GetInt(2);
        double scale = arguments.GetDouble(3, 1.);
        int offset_x = arguments.GetInt(4, 0), offset_y = arguments.GetInt(5, 0);
        auto pos = static_cast<libxl::Position>(arguments.GetInt(6, libxl::POSITION_MOVE_AND_SIZE));
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setPicture(row, col, id, scale, offset_x, offset_y, pos);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::SetPicture2) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0), col = arguments.GetInt(1), id = arguments.GetInt(2),
            width = arguments.GetInt(3, -1), height = arguments.GetInt(4, -1),
            offset_x = arguments.GetInt(5, 0), offset_y = arguments.GetInt(6, 0);
        auto pos = static_cast<libxl::Position>(arguments.GetInt(7, libxl::POSITION_MOVE_AND_SIZE));

        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setPicture2(row, col, id, width, height, offset_x, offset_y, pos);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::GetHorPageBreak) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int index = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->getHorPageBreak(index)));
    }

    NAN_METHOD(Sheet::GetHorPageBreakSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->getHorPageBreakSize()));
    }

    NAN_METHOD(Sheet::GetVerPageBreak) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int index = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->getVerPageBreak(index)));
    }

    NAN_METHOD(Sheet::GetVerPageBreakSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->getVerPageBreakSize()));
    }

    NAN_METHOD(Sheet::SetHorPageBreak) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        bool pagebreak = arguments.GetBoolean(1, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

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

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->setVerPageBreak(col, pagebreak)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::Split) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0), col = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->split(row, col);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::SplitInfo) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int row, col;

        if (!that->GetWrapped()->splitInfo(&row, &col)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("row").ToLocalChecked(), Nan::New<Integer>(row));
        Nan::Set(result, Nan::New<String>("col").ToLocalChecked(), Nan::New<Integer>(col));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::GroupRows) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int rowFirst = arguments.GetInt(0), rowLast = arguments.GetInt(1);
        bool collapsed = arguments.GetBoolean(2, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->groupRows(rowFirst, rowLast, collapsed)) {
            return util::ThrowLibxlError(that);
        };

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::GroupCols) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int colFirst = arguments.GetInt(0), colLast = arguments.GetInt(1);
        bool collapsed = arguments.GetBoolean(2, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->groupCols(colFirst, colLast, collapsed)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::GroupSummaryBelow) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->groupSummaryBelow()));
    }

    NAN_METHOD(Sheet::SetGroupSummaryBelow) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool summaryBelow = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setGroupSummaryBelow(summaryBelow);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::GroupSummaryRight) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->groupSummaryRight()));
    }

    NAN_METHOD(Sheet::SetGroupSummaryRight) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool summaryRight = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setGroupSummaryRight(summaryRight);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::Clear) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int rowFirst = arguments.GetInt(0, 0), rowLast = arguments.GetInt(1, 1048575),
            colFirst = arguments.GetInt(2, 0), colLast = arguments.GetInt(3, 16383);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->clear(rowFirst, rowLast, colFirst, colLast);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::InsertRow) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int rowFirst = arguments.GetInt(0), rowLast = arguments.GetInt(1);
        bool updateNamedRanges = arguments.GetBoolean(2, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->insertRow(rowFirst, rowLast, updateNamedRanges)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::InsertRowAsync) {
        class Worker : public AsyncWorker<Sheet> {
           public:
            Worker(Nan::Callback* callback, Local<Object> that, int rowFirst, int rowLast,
                   bool updateNamedRanges)
                : AsyncWorker<Sheet>(callback, that, "node-libxl-sheet-insert-row"),
                  rowFirst(rowFirst),
                  rowLast(rowLast),
                  updateNamedRanges(updateNamedRanges) {}

            virtual void Execute() {
                if (!that->GetWrapped()->insertRow(rowFirst, rowLast, updateNamedRanges)) {
                    RaiseLibxlError();
                }
            }

           private:
            int rowFirst, rowLast;
            bool updateNamedRanges;
        };

        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        if (arguments.Length() > 4) {
            return Nan::ThrowError("too many arguments");
        }

        int rowFirst = arguments.GetInt(0), rowLast = arguments.GetInt(1);
        bool updateNamedRanges = arguments.Length() > 3 ? arguments.GetBoolean(2, true) : true;
        Local<Function> callback = arguments.GetFunction(arguments.Length() - 1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        Nan::AsyncQueueWorker(new Worker(new Nan::Callback(callback), info.This(), rowFirst,
                                         rowLast, updateNamedRanges));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::InsertCol) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int colFirst = arguments.GetInt(0), colLast = arguments.GetInt(1);
        bool updateNamedRanges = arguments.GetBoolean(2, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->insertCol(colFirst, colLast, updateNamedRanges)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::InsertColAsync) {
        class Worker : public AsyncWorker<Sheet> {
           public:
            Worker(Nan::Callback* callback, Local<Object> that, int colFirst, int colLast,
                   bool updateNamedRanges)
                : AsyncWorker<Sheet>(callback, that, "node-libxl-insert-col"),
                  colFirst(colFirst),
                  colLast(colLast),
                  updateNamedRanges(updateNamedRanges) {}

            virtual void Execute() {
                if (!that->GetWrapped()->insertCol(colFirst, colLast, updateNamedRanges)) {
                    RaiseLibxlError();
                }
            }

           private:
            int colFirst, colLast;
            bool updateNamedRanges;
        };

        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        if (arguments.Length() > 4) {
            return Nan::ThrowError("too many arguments");
        }

        int colFirst = arguments.GetInt(0), colLast = arguments.GetInt(1);
        bool updateNamedRanges = arguments.Length() > 3 ? arguments.GetBoolean(2, true) : true;
        Local<Function> callback = arguments.GetFunction(arguments.Length() - 1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        Nan::AsyncQueueWorker(new Worker(new Nan::Callback(callback), info.This(), colFirst,
                                         colLast, updateNamedRanges));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::RemoveRow) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int rowFirst = arguments.GetInt(0), rowLast = arguments.GetInt(1);
        bool updateNamedRanges = arguments.GetBoolean(2, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->removeRow(rowFirst, rowLast, updateNamedRanges)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::RemoveRowAsync) {
        class Worker : public AsyncWorker<Sheet> {
           public:
            Worker(Nan::Callback* callback, Local<Object> that, int rowFirst, int rowLast,
                   bool updateNamedRanges)
                : AsyncWorker<Sheet>(callback, that, "node-livxl-remove-row"),
                  rowFirst(rowFirst),
                  rowLast(rowLast),
                  updateNamedRanges(updateNamedRanges) {}

            virtual void Execute() {
                if (!that->GetWrapped()->removeRow(rowFirst, rowLast, updateNamedRanges)) {
                    RaiseLibxlError();
                }
            }

           private:
            int rowFirst, rowLast;
            bool updateNamedRanges;
        };

        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        if (arguments.Length() > 4) {
            return Nan::ThrowError("too many arguments");
        }

        int rowFirst = arguments.GetInt(0), rowLast = arguments.GetInt(1);
        bool updateNamedRanges = arguments.Length() > 3 ? arguments.GetBoolean(2, true) : true;
        Local<Function> callback = arguments.GetFunction(arguments.Length() - 1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        Nan::AsyncQueueWorker(new Worker(new Nan::Callback(callback), info.This(), rowFirst,
                                         rowLast, updateNamedRanges));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::RemoveColAsync) {
        class Worker : public AsyncWorker<Sheet> {
           public:
            Worker(Nan::Callback* callback, Local<Object> that, int colFirst, int colLast,
                   bool removeNamedRanges)
                : AsyncWorker<Sheet>(callback, that, "node-libxl-remove-col"),
                  colFirst(colFirst),
                  colLast(colLast),
                  removeNamedRanges(removeNamedRanges) {}

            virtual void Execute() {
                if (!that->GetWrapped()->removeCol(colFirst, colLast, removeNamedRanges)) {
                    RaiseLibxlError();
                }
            }

           private:
            int colFirst, colLast;
            bool removeNamedRanges;
        };

        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        if (arguments.Length() > 4) {
            return Nan::ThrowError("too many arguments");
        }

        int colFirst = arguments.GetInt(0), colLast = arguments.GetInt(1);
        bool updateNamedRanges = arguments.Length() > 3 ? arguments.GetBoolean(2, true) : true;
        Local<Function> callback = arguments.GetFunction(arguments.Length() - 1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        Nan::AsyncQueueWorker(new Worker(new Nan::Callback(callback), info.This(), colFirst,
                                         colLast, updateNamedRanges));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::RemoveCol) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int colFirst = arguments.GetInt(0), colLast = arguments.GetInt(1);
        bool updateNamedRanges = arguments.GetBoolean(2, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->removeCol(colFirst, colLast, updateNamedRanges)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::CopyCell) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int rowSrc = arguments.GetInt(0), colSrc = arguments.GetInt(1),
            rowDst = arguments.GetInt(2), colDst = arguments.GetInt(3);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->copyCell(rowSrc, colSrc, rowDst, colDst)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::FirstRow) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->firstRow()));
    }

    NAN_METHOD(Sheet::LastRow) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->lastRow()));
    }

    NAN_METHOD(Sheet::FirstCol) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->firstCol()));
    }

    NAN_METHOD(Sheet::LastCol) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->lastCol()));
    }

    NAN_METHOD(Sheet::FirstFilledRow) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->firstFilledRow()));
    }

    NAN_METHOD(Sheet::LastFilledRow) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->lastFilledRow()));
    }

    NAN_METHOD(Sheet::FirstFilledCol) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->firstFilledCol()));
    }

    NAN_METHOD(Sheet::LastFilledCol) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->lastFilledCol()));
    }

    NAN_METHOD(Sheet::DisplayGridlines) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->displayGridlines()));
    }

    NAN_METHOD(Sheet::SetDisplayGridlines) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool displayGridlines = arguments.GetBoolean(0, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setDisplayGridlines(displayGridlines);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::PrintGridlines) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->printGridlines()));
    }

    NAN_METHOD(Sheet::SetPrintGridlines) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool printGridlines = arguments.GetBoolean(0, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setPrintGridlines(printGridlines);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::Zoom) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->zoom()));
    }

    NAN_METHOD(Sheet::SetZoom) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int zoom = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setZoom(zoom);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::PrintZoom) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->printZoom()));
    }

    NAN_METHOD(Sheet::SetPrintZoom) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int zoom = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setPrintZoom(zoom);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::GetPrintFit) {
        Nan::HandleScope scope;

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int wPages, hPages;

        if (that->GetWrapped()->getPrintFit(&wPages, &hPages)) {
            Local<Object> result = Nan::New<Object>();

            Nan::Set(result, Nan::New<String>("wPages").ToLocalChecked(),
                     Nan::New<Integer>(wPages));
            Nan::Set(result, Nan::New<String>("hPages").ToLocalChecked(),
                     Nan::New<Integer>(hPages));

            info.GetReturnValue().Set(result);
        } else {
            info.GetReturnValue().Set(Nan::False());
        }
    }

    NAN_METHOD(Sheet::SetPrintFit) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int wPages = arguments.GetInt(0, 1), hPages = arguments.GetInt(1, 1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setPrintFit(wPages, hPages);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::Landscape) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->landscape()));
    }

    NAN_METHOD(Sheet::SetLandscape) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool landscape = arguments.GetBoolean(0, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setLandscape(landscape);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::Paper) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->paper()));
    }

    NAN_METHOD(Sheet::SetPaper) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int paper = arguments.GetInt(0, libxl::PAPER_DEFAULT);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setPaper(static_cast<libxl::Paper>(paper));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::Header) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<String>(that->GetWrapped()->header()).ToLocalChecked());
    }

    NAN_METHOD(Sheet::SetHeader) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(header, arguments.GetString(0));
        double margin = arguments.GetDouble(1, 0.5);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->setHeader(*header, margin)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::HeaderMargin) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->headerMargin()));
    }

    NAN_METHOD(Sheet::Footer) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<String>(that->GetWrapped()->footer()).ToLocalChecked());
    }

    NAN_METHOD(Sheet::SetFooter) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(footer, arguments.GetString(0));
        double margin = arguments.GetDouble(1, 0.5);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->setFooter(*footer, margin)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::FooterMargin) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->footerMargin()));
    }

    NAN_METHOD(Sheet::HCenter) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->hCenter()));
    }

    NAN_METHOD(Sheet::SetHCenter) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool center = arguments.GetBoolean(0, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setHCenter(center);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::VCenter) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->vCenter()));
    }

    NAN_METHOD(Sheet::SetVCenter) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool center = arguments.GetBoolean(0, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setVCenter(center);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::MarginLeft) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->marginLeft()));
    }

    NAN_METHOD(Sheet::SetMarginLeft) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        double margin = arguments.GetDouble(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setMarginLeft(margin);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::MarginRight) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->marginRight()));
    }

    NAN_METHOD(Sheet::SetMarginRight) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        double margin = arguments.GetDouble(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setMarginRight(margin);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::MarginTop) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->marginTop()));
    }

    NAN_METHOD(Sheet::SetMarginTop) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        double margin = arguments.GetDouble(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setMarginTop(margin);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::MarginBottom) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->marginBottom()));
    }

    NAN_METHOD(Sheet::SetMarginBottom) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        double margin = arguments.GetDouble(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setMarginBottom(margin);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::PrintRowCol) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->printRowCol()));
    }

    NAN_METHOD(Sheet::SetPrintRowCol) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool printRowCol = arguments.GetBoolean(0, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setPrintRowCol(printRowCol);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::PrintRepeatRows) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int rowFirst, rowLast;

        if (!that->GetWrapped()->printRepeatRows(&rowFirst, &rowLast)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("rowFirst").ToLocalChecked(),
                 Nan::New<Integer>(rowFirst));
        Nan::Set(result, Nan::New<String>("rowLast").ToLocalChecked(), Nan::New<Integer>(rowLast));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::SetPrintRepeatRows) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int rowFirst = arguments.GetInt(0), rowLast = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setPrintRepeatRows(rowFirst, rowLast);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::PrintRepeatCols) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int colFirst, colLast;

        if (!that->GetWrapped()->printRepeatCols(&colFirst, &colLast)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("colFirst").ToLocalChecked(),
                 Nan::New<Integer>(colFirst));
        Nan::Set(result, Nan::New<String>("colLast").ToLocalChecked(), Nan::New<Integer>(colLast));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::SetPrintRepeatCols) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int colFirst = arguments.GetInt(0), colLast = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setPrintRepeatCols(colFirst, colLast);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::SetPrintArea) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int rowFirst = arguments.GetInt(0), rowLast = arguments.GetInt(1),
            colFirst = arguments.GetInt(2), colLast = arguments.GetInt(3);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setPrintArea(rowFirst, rowLast, colFirst, colLast);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::ClearPrintRepeats) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->clearPrintRepeats();

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::ClearPrintArea) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->clearPrintArea();

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::GetNamedRange) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(name, arguments.GetString(0));
        int scopeId = arguments.GetInt(1, libxl::SCOPE_UNDEFINED);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int rowFirst, rowLast, colFirst, colLast;
        bool hidden;

        if (!that->GetWrapped()->getNamedRange(*name, &rowFirst, &rowLast, &colFirst, &colLast,
                                               scopeId, &hidden)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();

        Nan::Set(result, Nan::New<String>("rowFirst").ToLocalChecked(),
                 Nan::New<Integer>(rowFirst));
        Nan::Set(result, Nan::New<String>("rowLast").ToLocalChecked(), Nan::New<Integer>(rowLast));
        Nan::Set(result, Nan::New<String>("colFirst").ToLocalChecked(),
                 Nan::New<Integer>(colFirst));
        Nan::Set(result, Nan::New<String>("colLast").ToLocalChecked(), Nan::New<Integer>(colLast));
        Nan::Set(result, Nan::New<String>("hidden").ToLocalChecked(), Nan::New<Boolean>(hidden));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::SetNamedRange) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(name, arguments.GetString(0));
        int rowFirst = arguments.GetInt(1), rowLast = arguments.GetInt(2),
            colFirst = arguments.GetInt(3), colLast = arguments.GetInt(4),
            scopeId = arguments.GetInt(5, libxl::SCOPE_UNDEFINED);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->setNamedRange(*name, rowFirst, rowLast, colFirst, colLast,
                                               static_cast<libxl::Scope>(scopeId))) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::GetTable) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(name, arguments.GetString(0));

        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int rowFirst, rowLast, colFirst, colLast, headerRowCount, totalRowCount;
        if (!that->GetWrapped()->getTable(*name, &rowFirst, &rowLast, &colFirst, &colLast,
                                          &headerRowCount, &totalRowCount)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("rowFirst").ToLocalChecked(),
                 Nan::New<Integer>(rowFirst));
        Nan::Set(result, Nan::New<String>("rowLast").ToLocalChecked(), Nan::New<Integer>(rowLast));
        Nan::Set(result, Nan::New<String>("colFirst").ToLocalChecked(),
                 Nan::New<Integer>(colFirst));
        Nan::Set(result, Nan::New<String>("colLast").ToLocalChecked(), Nan::New<Integer>(colLast));
        Nan::Set(result, Nan::New<String>("rowFirst").ToLocalChecked(),
                 Nan::New<Integer>(rowFirst));
        Nan::Set(result, Nan::New<String>("totalRowCount").ToLocalChecked(),
                 Nan::New<Integer>(headerRowCount));
        Nan::Set(result, Nan::New<String>("totalRowCount").ToLocalChecked(),
                 Nan::New<Integer>(rowFirst));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::TableSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->tableSize()));
    }

    NAN_METHOD(Sheet::Table) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int index = arguments.GetInt(0);

        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int rowFirst, rowLast, colFirst, colLast, headerRowCount, totalRowCount;
        if (!that->GetWrapped()->table(index, &rowFirst, &rowLast, &colFirst, &colLast,
                                       &headerRowCount, &totalRowCount)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("rowFirst").ToLocalChecked(),
                 Nan::New<Integer>(rowFirst));
        Nan::Set(result, Nan::New<String>("rowLast").ToLocalChecked(), Nan::New<Integer>(rowLast));
        Nan::Set(result, Nan::New<String>("colFirst").ToLocalChecked(),
                 Nan::New<Integer>(colFirst));
        Nan::Set(result, Nan::New<String>("colLast").ToLocalChecked(), Nan::New<Integer>(colLast));
        Nan::Set(result, Nan::New<String>("rowFirst").ToLocalChecked(),
                 Nan::New<Integer>(rowFirst));
        Nan::Set(result, Nan::New<String>("totalRowCount").ToLocalChecked(),
                 Nan::New<Integer>(headerRowCount));
        Nan::Set(result, Nan::New<String>("totalRowCount").ToLocalChecked(),
                 Nan::New<Integer>(rowFirst));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::HyperlinkSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->hyperlinkSize()));
    }

    NAN_METHOD(Sheet::Hyperlink) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int index = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int rowFirst, rowLast, colFirst, colLast;
        const char* hyperlink =
            that->GetWrapped()->hyperlink(index, &rowFirst, &rowLast, &colFirst, &colLast);

        if (!hyperlink) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("hyperlink").ToLocalChecked(),
                 Nan::New<String>(hyperlink).ToLocalChecked());
        Nan::Set(result, Nan::New<String>("rowFirst").ToLocalChecked(),
                 Nan::New<Integer>(rowFirst));
        Nan::Set(result, Nan::New<String>("rowLast").ToLocalChecked(), Nan::New<Integer>(rowLast));
        Nan::Set(result, Nan::New<String>("colFirst").ToLocalChecked(),
                 Nan::New<Integer>(colFirst));
        Nan::Set(result, Nan::New<String>("colLast").ToLocalChecked(), Nan::New<Integer>(colLast));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::DelHyperlink) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int index = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->delHyperlink(index)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::AddHyperlink) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(hyperlink, arguments.GetString(0));
        int rowFirst = arguments.GetInt(1);
        int rowLast = arguments.GetInt(2);
        int colFirst = arguments.GetInt(3);
        int colLast = arguments.GetInt(4);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->addHyperlink(*hyperlink, rowFirst, rowLast, colFirst, colLast);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::HyperlinkIndex) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int index = that->GetWrapped()->hyperlinkIndex(row, col);

        info.GetReturnValue().Set(Nan::New<Integer>(index));
    }

    NAN_METHOD(Sheet::IsAutoFilter) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->isAutoFilter()));
    }

    NAN_METHOD(Sheet::AutoFilter) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        libxl::AutoFilter* autoFilter = that->GetWrapped()->autoFilter();
        if (!autoFilter) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(AutoFilter::NewInstance(autoFilter, that->GetBookHandle()));
    }

    NAN_METHOD(Sheet::ApplyFilter) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->applyFilter();

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::RemoveFilter) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->removeFilter();

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::DelNamedRange) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(name, arguments.GetString(0));
        int scopeId = arguments.GetInt(1, libxl::SCOPE_UNDEFINED);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->delNamedRange(*name, static_cast<libxl::Scope>(scopeId))) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::NamedRangeSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->namedRangeSize()));
    }

    NAN_METHOD(Sheet::NamedRange) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int index = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int rowFirst, rowLast, colFirst, colLast, scopeId;
        bool hidden;
        const char* name = that->GetWrapped()->namedRange(index, &rowFirst, &rowLast, &colFirst,
                                                          &colLast, &scopeId, &hidden);

        if (!name) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();

        Nan::Set(result, Nan::New<String>("name").ToLocalChecked(),
                 Nan::New<String>(name).ToLocalChecked());
        Nan::Set(result, Nan::New<String>("rowFirst").ToLocalChecked(),
                 Nan::New<Integer>(rowFirst));
        Nan::Set(result, Nan::New<String>("rowLast").ToLocalChecked(), Nan::New<Integer>(rowLast));
        Nan::Set(result, Nan::New<String>("colFirst").ToLocalChecked(),
                 Nan::New<Integer>(colFirst));
        Nan::Set(result, Nan::New<String>("colLast").ToLocalChecked(), Nan::New<Integer>(colLast));
        Nan::Set(result, Nan::New<String>("scopeId").ToLocalChecked(), Nan::New<Integer>(scopeId));
        Nan::Set(result, Nan::New<String>("hidden").ToLocalChecked(), Nan::New<Boolean>(hidden));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::Name) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<String>(that->GetWrapped()->name()).ToLocalChecked());
    }

    NAN_METHOD(Sheet::SetName) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(name, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setName(*name);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::Protect) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->protect()));
    }

    NAN_METHOD(Sheet::SetProtect) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool protect = arguments.GetBoolean(0, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setProtect(protect);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::RightToLeft) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->rightToLeft()));
    }

    NAN_METHOD(Sheet::SetRightToLeft) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool rightToLeft = arguments.GetBoolean(0, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setRightToLeft(rightToLeft);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::Hidden) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->hidden()));
    }

    NAN_METHOD(Sheet::SetHidden) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int state = arguments.GetInt(0, libxl::SHEETSTATE_HIDDEN);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->setHidden(static_cast<libxl::SheetState>(state))) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::GetTopLeftView) {
        Nan::HandleScope scope;

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int row, col;
        that->GetWrapped()->getTopLeftView(&row, &col);

        Local<Object> result = Nan::New<Object>();

        Nan::Set(result, Nan::New<String>("row").ToLocalChecked(), Nan::New<Integer>(row));
        Nan::Set(result, Nan::New<String>("col").ToLocalChecked(), Nan::New<Integer>(col));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::SetTopLeftView) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0), col = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setTopLeftView(row, col);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::AddrToRowCol) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(addr, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int row = -1, col = -1;
        bool rowRelative = true, colRelative = true;

        that->GetWrapped()->addrToRowCol(*addr, &row, &col, &rowRelative, &colRelative);

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("row").ToLocalChecked(), Nan::New<Integer>(row));
        Nan::Set(result, Nan::New<String>("col").ToLocalChecked(), Nan::New<Integer>(col));
        Nan::Set(result, Nan::New<String>("rowRelative").ToLocalChecked(),
                 Nan::New<Boolean>(rowRelative));
        Nan::Set(result, Nan::New<String>("colRelative").ToLocalChecked(),
                 Nan::New<Boolean>(colRelative));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::RowColToAddr) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0), col = arguments.GetInt(1);
        bool rowRelative = arguments.GetBoolean(2, true),
             colRelative = arguments.GetBoolean(3, true);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        const char* addr = that->GetWrapped()->rowColToAddr(row, col, rowRelative, colRelative);

        if (!addr) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(addr).ToLocalChecked());
    }

    NAN_METHOD(Sheet::SetAutoFitArea) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int rowFirst = arguments.GetInt(0, 0);
        int colFirst = arguments.GetInt(1, 0);
        int rowLast = arguments.GetInt(2, -1);
        int colLast = arguments.GetInt(3, -1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setAutoFitArea(rowFirst, colFirst, rowLast, colLast);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::TabColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        libxl::Color color = that->GetWrapped()->tabColor();

        info.GetReturnValue().Set(Nan::New<Integer>(color));
    }

    NAN_METHOD(Sheet::SetTabColor) {
        Nan::HandleScope scope;

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        ArgumentHelper arguments(info);

        switch (arguments.Length()) {
            case 1: {
                int color = arguments.GetInt(0);
                ASSERT_ARGUMENTS(arguments);

                that->GetWrapped()->setTabColor(static_cast<libxl::Color>(color));

                break;
            }

            case 3: {
                int red = arguments.GetInt(0);
                int green = arguments.GetInt(1);
                int blue = arguments.GetInt(2);
                ASSERT_ARGUMENTS(arguments);

                that->GetWrapped()->setTabColor(red, green, blue);

                break;
            }

            default:
                return Nan::ThrowError("Invalid number of arguments");
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::GetTabColor) {
        Nan::HandleScope scope;

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int red, green, blue;
        if (!that->GetWrapped()->getTabColor(&red, &green, &blue)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("red").ToLocalChecked(), Nan::New<Integer>(red));
        Nan::Set(result, Nan::New<String>("green").ToLocalChecked(), Nan::New<Integer>(green));
        Nan::Set(result, Nan::New<String>("blue").ToLocalChecked(), Nan::New<Integer>(blue));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::AddIgnoredError) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int rowFirst = arguments.GetInt(0);
        int colFirst = arguments.GetInt(1);
        int rowLast = arguments.GetInt(2);
        int colLast = arguments.GetInt(3);
        int iError = arguments.GetInt(4);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        if (!that->GetWrapped()->addIgnoredError(rowFirst, colFirst, rowLast, colLast,
                                                 static_cast<libxl::IgnoredError>(iError))) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    /*
            virtual             void XLAPIENTRY addDataValidation(DataValidationType type,
       DataValidationOperator op, int rowFirst, int rowLast, int colFirst, int colLast, const TCHAR*
       value1, const TCHAR* value2 = 0, bool allowBlank = true, bool hideDropDown = false, bool
       showInputMessage = true, bool showErrorMessage = true, const TCHAR* promptTitle = 0, const
       TCHAR* prompt = 0, const TCHAR* errorTitle = 0, const TCHAR* error = 0,
       DataValidationErrorStyle errorStyle = VALIDATION_ERRSTYLE_STOP) = 0;

    */

    NAN_METHOD(Sheet::AddDataValidation) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        auto type = static_cast<libxl::DataValidationType>(arguments.GetInt(0));
        auto op = static_cast<libxl::DataValidationOperator>(arguments.GetInt(1));
        int rowFirst = arguments.GetInt(2);
        int rowLast = arguments.GetInt(3);
        int colFirst = arguments.GetInt(4);
        int colLast = arguments.GetInt(5);
        CSNanUtf8Value(value1, arguments.GetString(6));
        CSNanUtf8Value(value2, arguments.GetString(7, ""));
        bool allowBlank = arguments.GetBoolean(8, true);
        bool hideDropDown = arguments.GetBoolean(9, false);
        bool showInputMessage = arguments.GetBoolean(10, true);
        bool showErrorMessage = arguments.GetBoolean(11, true);
        CSNanUtf8Value(promptTitle, arguments.GetString(12, ""));
        CSNanUtf8Value(prompt, arguments.GetString(13, ""));
        CSNanUtf8Value(errorTitle, arguments.GetString(14, ""));
        CSNanUtf8Value(error, arguments.GetString(15, ""));
        auto errorStyle = static_cast<libxl::DataValidationErrorStyle>(
            arguments.GetInt(16, libxl::VALIDATION_ERRSTYLE_STOP));
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->addDataValidation(type, op, rowFirst, rowLast, colFirst, colLast,
                                              *value1, *value2, allowBlank, hideDropDown,
                                              showInputMessage, showErrorMessage, *promptTitle,
                                              *prompt, *errorTitle, *error, errorStyle);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::AddDataValidationDouble) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        auto type = static_cast<libxl::DataValidationType>(arguments.GetInt(0));
        auto op = static_cast<libxl::DataValidationOperator>(arguments.GetInt(1));
        int rowFirst = arguments.GetInt(2);
        int rowLast = arguments.GetInt(3);
        int colFirst = arguments.GetInt(4);
        int colLast = arguments.GetInt(5);
        double value1 = arguments.GetDouble(6);
        double value2 = arguments.GetDouble(7);
        bool allowBlank = arguments.GetBoolean(8, true);
        bool hideDropDown = arguments.GetBoolean(9, false);
        bool showInputMessage = arguments.GetBoolean(10, true);
        bool showErrorMessage = arguments.GetBoolean(11, true);
        CSNanUtf8Value(promptTitle, arguments.GetString(12, ""));
        CSNanUtf8Value(prompt, arguments.GetString(13, ""));
        CSNanUtf8Value(errorTitle, arguments.GetString(14, ""));
        CSNanUtf8Value(error, arguments.GetString(15, ""));
        auto errorStyle = static_cast<libxl::DataValidationErrorStyle>(
            arguments.GetInt(16, libxl::VALIDATION_ERRSTYLE_STOP));
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->addDataValidationDouble(
            type, op, rowFirst, rowLast, colFirst, colLast, value1, value2, allowBlank,
            hideDropDown, showInputMessage, showErrorMessage, *promptTitle, *prompt, *errorTitle,
            *error, errorStyle);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::RemoveDataValidations) {
        Nan::HandleScope scope;

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->removeDataValidations();

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::FormControlSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->formControlSize()));
    }

    NAN_METHOD(Sheet::FormControl) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int index = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        libxl::FormControl* formControl = that->GetWrapped()->formControl(index);
        if (!formControl) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(FormControl::NewInstance(formControl, that->GetBookHandle()));
    }

    NAN_METHOD(Sheet::AddConditionalFormatting) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        libxl::ConditionalFormatting* conditionalFormatting =
            that->GetWrapped()->addConditionalFormatting();
        if (!conditionalFormatting) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(
            ConditionalFormatting::NewInstance(conditionalFormatting, that->GetBookHandle()));
    }

    NAN_METHOD(Sheet::GetActiveCell) {
        Nan::HandleScope scope;

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        int row, col;
        that->GetWrapped()->getActiveCell(&row, &col);

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("row").ToLocalChecked(), Nan::New<Integer>(row));
        Nan::Set(result, Nan::New<String>("col").ToLocalChecked(), Nan::New<Integer>(col));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(Sheet::SetActiveCell) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int row = arguments.GetInt(0);
        int col = arguments.GetInt(1);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->setActiveCell(row, col);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::SelectionRange) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        const char* result = that->GetWrapped()->selectionRange();
        if (!result) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(result).ToLocalChecked());
    }

    NAN_METHOD(Sheet::AddSelectionRange) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(sqref, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->addSelectionRange(*sqref);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Sheet::RemoveSelection) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Sheet* that = FromJS(info.This());
        ASSERT_SHEET(that);

        that->GetWrapped()->removeSelection();

        info.GetReturnValue().Set(info.This());
    }

    // Init

    void Sheet::Initialize(Local<Object> exports) {
        using namespace libxl;

        Nan::HandleScope scope;

        Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(util::StubConstructor);
        t->SetClassName(Nan::New<String>("Sheet").ToLocalChecked());
        t->InstanceTemplate()->SetInternalFieldCount(1);

        BookHolder::Initialize<Sheet>(t);

        Nan::SetPrototypeMethod(t, "cellType", CellType);
        Nan::SetPrototypeMethod(t, "isFormula", IsFormula);
        Nan::SetPrototypeMethod(t, "cellFormat", CellFormat);
        Nan::SetPrototypeMethod(t, "setCellFormat", SetCellFormat);
        Nan::SetPrototypeMethod(t, "readStr", ReadStr);
        Nan::SetPrototypeMethod(t, "readString", ReadStr);
        Nan::SetPrototypeMethod(t, "readRichStr", ReadRichStr);
        Nan::SetPrototypeMethod(t, "writeRichStr", WriteRichStr);
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
        Nan::SetPrototypeMethod(t, "writeFormulaNum", WriteFormulaNum);
        Nan::SetPrototypeMethod(t, "writeFormulaStr", WriteFormulaStr);
        Nan::SetPrototypeMethod(t, "writeFormulaBool", WriteFormulaBool);
        Nan::SetPrototypeMethod(t, "readComment", ReadComment);
        Nan::SetPrototypeMethod(t, "writeComment", WriteComment);
        Nan::SetPrototypeMethod(t, "removeComment", RemoveComment);
        Nan::SetPrototypeMethod(t, "isDate", IsDate);
        Nan::SetPrototypeMethod(t, "isRichStr", IsRichStr);
        Nan::SetPrototypeMethod(t, "readError", ReadError);
        Nan::SetPrototypeMethod(t, "writeError", WriteError);
        Nan::SetPrototypeMethod(t, "colWidth", ColWidth);
        Nan::SetPrototypeMethod(t, "rowHeight", RowHeight);
        Nan::SetPrototypeMethod(t, "colWidthPx", ColWidthPx);
        Nan::SetPrototypeMethod(t, "rowHeightPx", RowHeightPx);
        Nan::SetPrototypeMethod(t, "setCol", SetCol);
        Nan::SetPrototypeMethod(t, "setRow", SetRow);
        Nan::SetPrototypeMethod(t, "setColPx", SetColPx);
        Nan::SetPrototypeMethod(t, "setRowPx", SetRowPx);
        Nan::SetPrototypeMethod(t, "rowFormat", RowFormat);
        Nan::SetPrototypeMethod(t, "colFormat", ColFormat);
        Nan::SetPrototypeMethod(t, "rowHidden", RowHidden);
        Nan::SetPrototypeMethod(t, "setRowHidden", SetRowHidden);
        Nan::SetPrototypeMethod(t, "colHidden", ColHidden);
        Nan::SetPrototypeMethod(t, "setColHidden", SetColHidden);
        Nan::SetPrototypeMethod(t, "defaultRowHeight", DefaultRowHeight);
        Nan::SetPrototypeMethod(t, "setDefaultRowHeight", SetDefaultRowHeight);
        Nan::SetPrototypeMethod(t, "getMerge", GetMerge);
        Nan::SetPrototypeMethod(t, "setMerge", SetMerge);
        Nan::SetPrototypeMethod(t, "delMerge", DelMerge);
        Nan::SetPrototypeMethod(t, "mergeSize", MergeSize);
        Nan::SetPrototypeMethod(t, "merge", Merge);
        Nan::SetPrototypeMethod(t, "delMergeByIndex", DelMergeByIndex);
        Nan::SetPrototypeMethod(t, "pictureSize", PictureSize);
        Nan::SetPrototypeMethod(t, "getPicture", GetPicture);
        Nan::SetPrototypeMethod(t, "removePicture", RemovePicture);
        Nan::SetPrototypeMethod(t, "removePictureByIndex", RemovePictureByIndex);
        Nan::SetPrototypeMethod(t, "setPicture", SetPicture);
        Nan::SetPrototypeMethod(t, "setPicture2", SetPicture2);
        Nan::SetPrototypeMethod(t, "getHorPageBreak", GetHorPageBreak);
        Nan::SetPrototypeMethod(t, "getHorPageBreakSize", GetHorPageBreakSize);
        Nan::SetPrototypeMethod(t, "getVerPageBreak", GetVerPageBreak);
        Nan::SetPrototypeMethod(t, "getVerPageBreakSize", GetVerPageBreakSize);
        Nan::SetPrototypeMethod(t, "setHorPageBreak", SetHorPageBreak);
        Nan::SetPrototypeMethod(t, "setVerPageBreak", SetVerPageBreak);
        Nan::SetPrototypeMethod(t, "split", Split);
        Nan::SetPrototypeMethod(t, "splitInfo", SplitInfo);
        Nan::SetPrototypeMethod(t, "groupRows", GroupRows);
        Nan::SetPrototypeMethod(t, "groupCols", GroupCols);
        Nan::SetPrototypeMethod(t, "groupSummaryBelow", GroupSummaryBelow);
        Nan::SetPrototypeMethod(t, "setGroupSummaryBelow", SetGroupSummaryBelow);
        Nan::SetPrototypeMethod(t, "groupSummaryRight", GroupSummaryRight);
        Nan::SetPrototypeMethod(t, "setGroupSummaryRight", SetGroupSummaryRight);
        Nan::SetPrototypeMethod(t, "clear", Clear);
        Nan::SetPrototypeMethod(t, "insertRow", InsertRow);
        Nan::SetPrototypeMethod(t, "insertRowSync", InsertRow);
        Nan::SetPrototypeMethod(t, "insertRowAsync", InsertRowAsync);
        Nan::SetPrototypeMethod(t, "insertCol", InsertCol);
        Nan::SetPrototypeMethod(t, "insertColSync", InsertCol);
        Nan::SetPrototypeMethod(t, "insertColAsync", InsertColAsync);
        Nan::SetPrototypeMethod(t, "removeRow", RemoveRow);
        Nan::SetPrototypeMethod(t, "removeRowSync", RemoveRow);
        Nan::SetPrototypeMethod(t, "removeRowAsync", RemoveRowAsync);
        Nan::SetPrototypeMethod(t, "removeCol", RemoveCol);
        Nan::SetPrototypeMethod(t, "removeColSync", RemoveCol);
        Nan::SetPrototypeMethod(t, "removeColAsync", RemoveColAsync);
        Nan::SetPrototypeMethod(t, "copyCell", CopyCell);
        Nan::SetPrototypeMethod(t, "firstRow", FirstRow);
        Nan::SetPrototypeMethod(t, "lastRow", LastRow);
        Nan::SetPrototypeMethod(t, "firstCol", FirstCol);
        Nan::SetPrototypeMethod(t, "lastCol", LastCol);
        Nan::SetPrototypeMethod(t, "firstFilledRow", FirstFilledRow);
        Nan::SetPrototypeMethod(t, "lastFilledRow", LastFilledRow);
        Nan::SetPrototypeMethod(t, "firstFilledCol", FirstFilledCol);
        Nan::SetPrototypeMethod(t, "lastFilledCol", LastFilledCol);
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
        Nan::SetPrototypeMethod(t, "printRepeatRows", PrintRepeatRows);
        Nan::SetPrototypeMethod(t, "setPrintRepeatRows", SetPrintRepeatRows);
        Nan::SetPrototypeMethod(t, "printRepeatCols", PrintRepeatCols);
        Nan::SetPrototypeMethod(t, "setPrintRepeatCols", SetPrintRepeatCols);
        Nan::SetPrototypeMethod(t, "setPrintArea", SetPrintArea);
        Nan::SetPrototypeMethod(t, "clearPrintRepeats", ClearPrintRepeats);
        Nan::SetPrototypeMethod(t, "clearPrintArea", ClearPrintArea);
        Nan::SetPrototypeMethod(t, "getNamedRange", GetNamedRange);
        Nan::SetPrototypeMethod(t, "setNamedRange", SetNamedRange);
        Nan::SetPrototypeMethod(t, "setNamedRange", SetNamedRange);
        Nan::SetPrototypeMethod(t, "delNamedRange", DelNamedRange);
        Nan::SetPrototypeMethod(t, "getTable", GetTable);
        Nan::SetPrototypeMethod(t, "tableSize", TableSize);
        Nan::SetPrototypeMethod(t, "table", Table);
        Nan::SetPrototypeMethod(t, "namedRangeSize", NamedRangeSize);
        Nan::SetPrototypeMethod(t, "namedRange", NamedRange);
        Nan::SetPrototypeMethod(t, "name", Name);
        Nan::SetPrototypeMethod(t, "setName", SetName);
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
        Nan::SetPrototypeMethod(t, "hyperlinkSize", HyperlinkSize);
        Nan::SetPrototypeMethod(t, "hyperlink", Hyperlink);
        Nan::SetPrototypeMethod(t, "delHyperlink", DelHyperlink);
        Nan::SetPrototypeMethod(t, "addHyperlink", AddHyperlink);
        Nan::SetPrototypeMethod(t, "hyperlinkIndex", HyperlinkIndex);
        Nan::SetPrototypeMethod(t, "isAutoFilter", IsAutoFilter);
        Nan::SetPrototypeMethod(t, "autoFilter", AutoFilter);
        Nan::SetPrototypeMethod(t, "applyFilter", ApplyFilter);
        Nan::SetPrototypeMethod(t, "removeFilter", RemoveFilter);
        Nan::SetPrototypeMethod(t, "setAutoFitArea", SetAutoFitArea);
        Nan::SetPrototypeMethod(t, "tabColor", TabColor);
        Nan::SetPrototypeMethod(t, "setTabColor", SetTabColor);
        Nan::SetPrototypeMethod(t, "getTabColor", GetTabColor);
        Nan::SetPrototypeMethod(t, "addIgnoredError", AddIgnoredError);
        Nan::SetPrototypeMethod(t, "addDataValidation", AddDataValidation);
        Nan::SetPrototypeMethod(t, "addDataValidationDouble", AddDataValidationDouble);
        Nan::SetPrototypeMethod(t, "removeDataValidations", RemoveDataValidations);
        Nan::SetPrototypeMethod(t, "formControlSize", FormControlSize);
        Nan::SetPrototypeMethod(t, "formControl", FormControl);
        Nan::SetPrototypeMethod(t, "addConditionalFormatting", AddConditionalFormatting);
        Nan::SetPrototypeMethod(t, "getActiveCell", GetActiveCell);
        Nan::SetPrototypeMethod(t, "setActiveCell", SetActiveCell);
        Nan::SetPrototypeMethod(t, "selectionRange", SelectionRange);
        Nan::SetPrototypeMethod(t, "addSelectionRange", AddSelectionRange);
        Nan::SetPrototypeMethod(t, "removeSelection", RemoveSelection);

        t->ReadOnlyPrototype();
        constructor.Reset(Nan::GetFunction(t).ToLocalChecked());
        Nan::Set(exports, Nan::New<String>("Sheet").ToLocalChecked(), Nan::New(constructor));
    }

}  // namespace node_libxl
