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

#include "auto_filter.h"

#include "argument_helper.h"
#include "assert.h"
#include "filter_column.h"
#include "util.h"

using namespace v8;

namespace node_libxl {
    // Lifecycle

    AutoFilter::AutoFilter(libxl::AutoFilter* autoFilter, Local<Value> book)
        : Wrapper<libxl::AutoFilter, AutoFilter>(autoFilter), BookHolder(book) {}

    Local<Object> AutoFilter::NewInstance(libxl::AutoFilter* libxlAutoFilter, Local<Value> book) {
        Nan::EscapableHandleScope scope;

        AutoFilter* autoFilter = new AutoFilter(libxlAutoFilter, book);

        Local<Object> that = util::CallStubConstructor(Nan::New(constructor)).As<Object>();

        autoFilter->Wrap(that);

        return scope.Escape(that);
    }

    // Wrappers

    NAN_METHOD(AutoFilter::GetRef) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int rowFirst, rowLast, colFirst, colLast;

        ASSERT_ARGUMENTS(arguments);

        AutoFilter* that = FromJS(info.This());
        ASSERT_THIS(that);

        if (!that->GetWrapped()->getRef(&rowFirst, &rowLast, &colFirst, &colLast)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("rowFirst").ToLocalChecked(), Nan::New<Number>(rowFirst));
        Nan::Set(result, Nan::New<String>("rowLast").ToLocalChecked(), Nan::New<Number>(rowLast));
        Nan::Set(result, Nan::New<String>("colFirst").ToLocalChecked(), Nan::New<Number>(colFirst));
        Nan::Set(result, Nan::New<String>("colLast").ToLocalChecked(), Nan::New<Number>(colLast));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(AutoFilter::SetRef) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int rowFirst = arguments.GetInt(0);
        int rowLast = arguments.GetInt(1);
        int colFirst = arguments.GetInt(2);
        int colLast = arguments.GetInt(3);

        ASSERT_ARGUMENTS(arguments);

        AutoFilter* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setRef(rowFirst, rowLast, colFirst, colLast);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(AutoFilter::Column) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int colId = arguments.GetInt(0);

        ASSERT_ARGUMENTS(arguments);

        AutoFilter* that = FromJS(info.This());
        ASSERT_THIS(that);

        libxl::IFilterColumnT<char>* column = that->GetWrapped()->column(colId);
        if (!column) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(FilterColumn::NewInstance(column, that->GetBookHandle()));
    }

    NAN_METHOD(AutoFilter::ColumnSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        AutoFilter* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->columnSize()));
    }

    NAN_METHOD(AutoFilter::ColumnByIndex) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int index = arguments.GetInt(0);

        ASSERT_ARGUMENTS(arguments);

        AutoFilter* that = FromJS(info.This());
        ASSERT_THIS(that);

        libxl::IFilterColumnT<char>* column = that->GetWrapped()->columnByIndex(index);
        if (!column) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(FilterColumn::NewInstance(column, that->GetBookHandle()));
    }

    NAN_METHOD(AutoFilter::GetSortRange) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int rowFirst, rowLast, colFirst, colLast;

        ASSERT_ARGUMENTS(arguments);

        AutoFilter* that = FromJS(info.This());
        ASSERT_THIS(that);

        if (!that->GetWrapped()->getSortRange(&rowFirst, &rowLast, &colFirst, &colLast)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("rowFirst").ToLocalChecked(), Nan::New<Number>(rowFirst));
        Nan::Set(result, Nan::New<String>("rowLast").ToLocalChecked(), Nan::New<Number>(rowLast));
        Nan::Set(result, Nan::New<String>("colFirst").ToLocalChecked(), Nan::New<Number>(colFirst));
        Nan::Set(result, Nan::New<String>("colLast").ToLocalChecked(), Nan::New<Number>(colLast));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(AutoFilter::SortLevels) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        AutoFilter* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->sortLevels()));
    }

    NAN_METHOD(AutoFilter::GetSort) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int level = arguments.GetInt(0, 0);

        ASSERT_ARGUMENTS(arguments);

        AutoFilter* that = FromJS(info.This());
        ASSERT_THIS(that);

        int columnIndex;
        bool descending;

        if (!that->GetWrapped()->getSort(&columnIndex, &descending, level)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("columnIndex").ToLocalChecked(),
                 Nan::New<Number>(columnIndex));
        Nan::Set(result, Nan::New<String>("descending").ToLocalChecked(),
                 Nan::New<Boolean>(descending));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(AutoFilter::SetSort) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int columnIndex = arguments.GetInt(0);
        bool descending = arguments.GetBoolean(1, false);

        ASSERT_ARGUMENTS(arguments);

        AutoFilter* that = FromJS(info.This());
        ASSERT_THIS(that);

        if (!that->GetWrapped()->setSort(columnIndex, descending)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(AutoFilter::AddSort) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int columnIndex = arguments.GetInt(0);
        bool descending = arguments.GetBoolean(1, false);

        ASSERT_ARGUMENTS(arguments);

        AutoFilter* that = FromJS(info.This());
        ASSERT_THIS(that);

        if (!that->GetWrapped()->addSort(columnIndex, descending)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    // Init

    void AutoFilter::Initialize(Local<Object> exports) {
        using namespace libxl;

        Nan::HandleScope scope;

        Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(util::StubConstructor);
        t->SetClassName(Nan::New<String>("AutoFilter").ToLocalChecked());
        t->InstanceTemplate()->SetInternalFieldCount(1);

        BookHolder::Initialize<AutoFilter>(t);

        Nan::SetPrototypeMethod(t, "getRef", GetRef);
        Nan::SetPrototypeMethod(t, "setRef", SetRef);
        Nan::SetPrototypeMethod(t, "column", Column);
        Nan::SetPrototypeMethod(t, "columnSize", ColumnSize);
        Nan::SetPrototypeMethod(t, "columnByIndex", ColumnByIndex);
        Nan::SetPrototypeMethod(t, "getSortRange", GetSortRange);
        Nan::SetPrototypeMethod(t, "sortLevels", SortLevels);
        Nan::SetPrototypeMethod(t, "getSort", GetSort);
        Nan::SetPrototypeMethod(t, "setSort", SetSort);
        Nan::SetPrototypeMethod(t, "addSort", AddSort);

        t->ReadOnlyPrototype();
        constructor.Reset(Nan::GetFunction(t).ToLocalChecked());
        Nan::Set(exports, Nan::New<String>("AutoFilter").ToLocalChecked(), Nan::New(constructor));
    }
}  // namespace node_libxl
