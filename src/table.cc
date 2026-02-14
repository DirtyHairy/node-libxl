/**
 * The MIT License (MIT)
 *
 * Copyright (c) 2025 Christian Speckner <cnspeckn@googlemail.com>
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

#include "table.h"

#include "argument_helper.h"
#include "assert.h"
#include "auto_filter.h"
#include "util.h"

using namespace v8;

namespace node_libxl {
    // Lifecycle

    Table::Table(libxl::Table* table, Local<Value> book)
        : Wrapper<libxl::Table, Table>(table), BookHolder(book) {}

    Local<Object> Table::NewInstance(libxl::Table* libxlTable, Local<Value> book) {
        Nan::EscapableHandleScope scope;

        Table* table = new Table(libxlTable, book);

        Local<Object> that = util::CallStubConstructor(Nan::New(constructor)).As<Object>();

        table->Wrap(that);

        return scope.Escape(that);
    }

    // Wrappers

    NAN_METHOD(Table::Name) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<String>(that->GetWrapped()->name()).ToLocalChecked());
    }

    NAN_METHOD(Table::SetName) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(name, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setName(*name);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Table::Ref) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<String>(that->GetWrapped()->ref()).ToLocalChecked());
    }

    NAN_METHOD(Table::SetRef) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(ref, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setRef(*ref);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Table::TableAutoFilter) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        libxl::AutoFilter* autoFilter = that->GetWrapped()->autoFilter();
        if (!autoFilter) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(AutoFilter::NewInstance(autoFilter, that->GetBookHandle()));
    }

    NAN_METHOD(Table::Style) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->style()));
    }

    NAN_METHOD(Table::SetStyle) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int style = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setStyle(static_cast<libxl::TableStyle>(style));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Table::ShowRowStripes) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->showRowStripes()));
    }

    NAN_METHOD(Table::SetShowRowStripes) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        bool showRowStripes = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setShowRowStripes(showRowStripes);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Table::ShowColumnStripes) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->showColumnStripes()));
    }

    NAN_METHOD(Table::SetShowColumnStripes) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        bool showColumnStripes = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setShowColumnStripes(showColumnStripes);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Table::ShowFirstColumn) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->showFirstColumn()));
    }

    NAN_METHOD(Table::SetShowFirstColumn) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        bool showFirstColumn = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setShowFirstColumn(showFirstColumn);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Table::ShowLastColumn) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->showLastColumn()));
    }

    NAN_METHOD(Table::SetShowLastColumn) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        bool showLastColumn = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setShowLastColumn(showLastColumn);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Table::ColumnSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->columnSize()));
    }

    NAN_METHOD(Table::ColumnName) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int index = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        const char* name = that->GetWrapped()->columnName(index);
        if (!name) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(name).ToLocalChecked());
    }

    NAN_METHOD(Table::SetColumnName) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int index = arguments.GetInt(0);
        CSNanUtf8Value(name, arguments.GetString(1));
        ASSERT_ARGUMENTS(arguments);

        Table* that = FromJS(info.This());
        ASSERT_THIS(that);

        if (!that->GetWrapped()->setColumnName(index, *name)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    // Init

    void Table::Initialize(Local<Object> exports) {
        using namespace libxl;

        Nan::HandleScope scope;

        Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(util::StubConstructor);
        t->SetClassName(Nan::New<String>("Table").ToLocalChecked());
        t->InstanceTemplate()->SetInternalFieldCount(1);

        BookHolder::Initialize<Table>(t);

        Nan::SetPrototypeMethod(t, "name", Name);
        Nan::SetPrototypeMethod(t, "setName", SetName);
        Nan::SetPrototypeMethod(t, "ref", Ref);
        Nan::SetPrototypeMethod(t, "setRef", SetRef);
        Nan::SetPrototypeMethod(t, "autoFilter", TableAutoFilter);
        Nan::SetPrototypeMethod(t, "style", Style);
        Nan::SetPrototypeMethod(t, "setStyle", SetStyle);
        Nan::SetPrototypeMethod(t, "showRowStripes", ShowRowStripes);
        Nan::SetPrototypeMethod(t, "setShowRowStripes", SetShowRowStripes);
        Nan::SetPrototypeMethod(t, "showColumnStripes", ShowColumnStripes);
        Nan::SetPrototypeMethod(t, "setShowColumnStripes", SetShowColumnStripes);
        Nan::SetPrototypeMethod(t, "showFirstColumn", ShowFirstColumn);
        Nan::SetPrototypeMethod(t, "setShowFirstColumn", SetShowFirstColumn);
        Nan::SetPrototypeMethod(t, "showLastColumn", ShowLastColumn);
        Nan::SetPrototypeMethod(t, "setShowLastColumn", SetShowLastColumn);
        Nan::SetPrototypeMethod(t, "columnSize", ColumnSize);
        Nan::SetPrototypeMethod(t, "columnName", ColumnName);
        Nan::SetPrototypeMethod(t, "setColumnName", SetColumnName);

        t->ReadOnlyPrototype();
        constructor.Reset(Nan::GetFunction(t).ToLocalChecked());
        Nan::Set(exports, Nan::New<String>("Table").ToLocalChecked(), Nan::New(constructor));
    }
}  // namespace node_libxl
