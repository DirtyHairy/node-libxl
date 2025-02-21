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

#include "filter_column.h"

#include "argument_helper.h"
#include "assert.h"
#include "util.h"

using namespace v8;

namespace node_libxl {
    // Lifecycle

    FilterColumn::FilterColumn(libxl::IFilterColumnT<char>* filterColumn, Local<Value> book)
        : Wrapper<libxl::IFilterColumnT<char>, FilterColumn>(filterColumn), BookHolder(book) {}

    Local<Object> FilterColumn::NewInstance(libxl::IFilterColumnT<char>* libxlFilterColumn,
                                            Local<Value> book) {
        Nan::EscapableHandleScope scope;

        FilterColumn* filterColumn = new FilterColumn(libxlFilterColumn, book);

        Local<Object> that = util::CallStubConstructor(Nan::New(constructor)).As<Object>();

        filterColumn->Wrap(that);

        return scope.Escape(that);
    }

    // Wrappers

    NAN_METHOD(FilterColumn::Index) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FilterColumn* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->index()));
    }

    NAN_METHOD(FilterColumn::FilterType) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FilterColumn* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->filterType()));
    }

    NAN_METHOD(FilterColumn::FilterSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FilterColumn* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->filterSize()));
    }

    NAN_METHOD(FilterColumn::Filter) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int index = arguments.GetInt(0);

        ASSERT_ARGUMENTS(arguments);

        FilterColumn* that = FromJS(info.This());
        ASSERT_THIS(that);

        const char* filter = that->GetWrapped()->filter(index);
        if (!filter) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(filter).ToLocalChecked());
    }

    NAN_METHOD(FilterColumn::AddFilter) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        CSNanUtf8Value(value, arguments.GetString(0));

        ASSERT_ARGUMENTS(arguments);

        FilterColumn* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->addFilter(*value);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FilterColumn::GetTop10) {
        Nan::HandleScope scope;

        FilterColumn* that = FromJS(info.This());
        ASSERT_THIS(that);

        double value;
        bool top, percent;

        if (!that->GetWrapped()->getTop10(&value, &top, &percent)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("value").ToLocalChecked(), Nan::New<Number>(value));
        Nan::Set(result, Nan::New<String>("top").ToLocalChecked(), Nan::New<Boolean>(top));
        Nan::Set(result, Nan::New<String>("percent").ToLocalChecked(), Nan::New<Boolean>(percent));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(FilterColumn::SetTop10) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        double value = arguments.GetDouble(0);
        bool top = arguments.GetBoolean(1, true);
        bool percent = arguments.GetBoolean(2, false);

        ASSERT_ARGUMENTS(arguments);

        FilterColumn* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setTop10(value, top, percent);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FilterColumn::GetCustomFilter) {
        Nan::HandleScope scope;

        FilterColumn* that = FromJS(info.This());
        ASSERT_THIS(that);

        libxl::Operator op1, op2;
        const char *v1, *v2;
        bool andOp;

        if (!that->GetWrapped()->getCustomFilter(&op1, &v1, &op2, &v2, &andOp)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("op1").ToLocalChecked(), Nan::New<Number>(op1));
        Nan::Set(result, Nan::New<String>("v1").ToLocalChecked(),
                 Nan::New<String>(v1).ToLocalChecked());
        Nan::Set(result, Nan::New<String>("op2").ToLocalChecked(), Nan::New<Number>(op2));
        Nan::Set(result, Nan::New<String>("v2").ToLocalChecked(),
                 Nan::New<String>(v2).ToLocalChecked());
        Nan::Set(result, Nan::New<String>("andOp").ToLocalChecked(), Nan::New<Boolean>(andOp));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(FilterColumn::SetCustomFilter) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        auto op1 = static_cast<libxl::Operator>(arguments.GetInt(0));
        CSNanUtf8Value(v1, arguments.GetString(1));
        auto op2 = static_cast<libxl::Operator>(arguments.GetInt(2, libxl::OPERATOR_EQUAL));
        CSNanUtf8Value(v2, arguments.GetString(3, nullptr));
        bool andOp = arguments.GetBoolean(4, false);

        ASSERT_ARGUMENTS(arguments);

        FilterColumn* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setCustomFilter(op1, *v1, op2, *v2, andOp);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FilterColumn::Clear) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FilterColumn* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->clear();

        info.GetReturnValue().Set(info.This());
    }

    // Init

    void FilterColumn::Initialize(Local<Object> exports) {
        using namespace libxl;

        Nan::HandleScope scope;

        Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(util::StubConstructor);
        t->SetClassName(Nan::New<String>("FilterColumn").ToLocalChecked());
        t->InstanceTemplate()->SetInternalFieldCount(1);

        BookHolder::Initialize<FilterColumn>(t);

        Nan::SetPrototypeMethod(t, "index", Index);
        Nan::SetPrototypeMethod(t, "filterType", FilterType);
        Nan::SetPrototypeMethod(t, "filterSize", FilterSize);
        Nan::SetPrototypeMethod(t, "getFilter", Filter);
        Nan::SetPrototypeMethod(t, "addFilter", AddFilter);
        Nan::SetPrototypeMethod(t, "getTop10", GetTop10);
        Nan::SetPrototypeMethod(t, "setTop10", SetTop10);
        Nan::SetPrototypeMethod(t, "getCustomFilter", GetCustomFilter);
        Nan::SetPrototypeMethod(t, "setCustomFilter", SetCustomFilter);
        Nan::SetPrototypeMethod(t, "clear", Clear);

        t->ReadOnlyPrototype();
        constructor.Reset(Nan::GetFunction(t).ToLocalChecked());
        Nan::Set(exports, Nan::New<String>("FilterColumn").ToLocalChecked(), Nan::New(constructor));

        NODE_DEFINE_CONSTANT(exports, FILTER_VALUE);
        NODE_DEFINE_CONSTANT(exports, FILTER_TOP10);
        NODE_DEFINE_CONSTANT(exports, FILTER_CUSTOM);
        NODE_DEFINE_CONSTANT(exports, FILTER_DYNAMIC);
        NODE_DEFINE_CONSTANT(exports, FILTER_COLOR);
        NODE_DEFINE_CONSTANT(exports, FILTER_ICON);
        NODE_DEFINE_CONSTANT(exports, FILTER_EXT);
        NODE_DEFINE_CONSTANT(exports, FILTER_NOT_SET);

        NODE_DEFINE_CONSTANT(exports, OPERATOR_EQUAL);
        NODE_DEFINE_CONSTANT(exports, OPERATOR_GREATER_THAN);
        NODE_DEFINE_CONSTANT(exports, OPERATOR_GREATER_THAN_OR_EQUAL);
        NODE_DEFINE_CONSTANT(exports, OPERATOR_LESS_THAN);
        NODE_DEFINE_CONSTANT(exports, OPERATOR_LESS_THAN_OR_EQUAL);
        NODE_DEFINE_CONSTANT(exports, OPERATOR_NOT_EQUAL);
    }
}  // namespace node_libxl
