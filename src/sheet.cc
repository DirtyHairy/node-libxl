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

    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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

    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->isFormula(row, col)));
}


NAN_METHOD(Sheet::CellFormat) {
    NanScope();

    ArgumentHelper arguments(args);

    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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

    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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
    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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

    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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
    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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

    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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
    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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

    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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
    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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

    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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
    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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

    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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
    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
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

    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
    String::Utf8Value value(arguments.GetString(2));
    String::Utf8Value author(arguments.GetString(3, ""));
    int32_t width = arguments.GetInt(4, 129);
    int32_t height = arguments.GetInt(5, 75);
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

    int32_t row = arguments.GetInt(0);
    int32_t col = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->readError(row, col)));
}


NAN_METHOD(Sheet::ColWidth) {
    NanScope();

    ArgumentHelper arguments(args);

    int32_t col = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Number>(that->GetWrapped()->colWidth(col)));
}


NAN_METHOD(Sheet::RowHeight) {
    NanScope();

    ArgumentHelper arguments(args);

    int32_t row = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Number>(that->GetWrapped()->rowHeight(row)));
}


NAN_METHOD(Sheet::SetCol) {
    NanScope();

    ArgumentHelper arguments(args);

    int32_t first = arguments.GetInt(0);
    int32_t last = arguments.GetInt(1);
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

    int32_t row = arguments.GetInt(0);
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

    int32_t row = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->rowHidden(row)));
}


NAN_METHOD(Sheet::SetRowHidden) {
    NanScope();

    ArgumentHelper arguments(args);

    int32_t row = arguments.GetInt(0);
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

    int32_t col = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->colHidden(col)));
}


NAN_METHOD(Sheet::SetColHidden) {
    NanScope();

    ArgumentHelper arguments(args);

    int32_t col = arguments.GetInt(0);
    bool hidden = arguments.GetBoolean(1);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setColHidden(col, hidden)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Sheet::SetMerge) {
    NanScope();

    ArgumentHelper arguments(args);

    int32_t rowFirst = arguments.GetInt(0);
    int32_t rowLast = arguments.GetInt(1);
    int32_t colFirst = arguments.GetInt(2);
    int32_t colLast = arguments.GetInt(3);
    ASSERT_ARGUMENTS(arguments);

    Sheet* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setMerge(rowFirst, rowLast, colFirst, colLast)) {
        return util::ThrowLibxlError(that);
    }
    
    NanReturnValue(args.This());
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
    NODE_SET_PROTOTYPE_METHOD(t, "readError", ReadError);
    NODE_SET_PROTOTYPE_METHOD(t, "colWidth", ColWidth);
    NODE_SET_PROTOTYPE_METHOD(t, "rowHeight", RowHeight);
    NODE_SET_PROTOTYPE_METHOD(t, "setCol", SetCol);
    NODE_SET_PROTOTYPE_METHOD(t, "setRow", SetRow);
    NODE_SET_PROTOTYPE_METHOD(t, "rowHidden", RowHidden);
    NODE_SET_PROTOTYPE_METHOD(t, "setRowHidden", SetRowHidden);
    NODE_SET_PROTOTYPE_METHOD(t, "colHidden", ColHidden);
    NODE_SET_PROTOTYPE_METHOD(t, "setColHidden", SetColHidden);
    NODE_SET_PROTOTYPE_METHOD(t, "setMerge", SetMerge);

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
}


}
