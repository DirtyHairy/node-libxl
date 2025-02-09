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

#include "core_properties.h"

#include "argument_helper.h"
#include "assert.h"
#include "font.h"
#include "util.h"

using namespace v8;

namespace node_libxl {
    // Lifecycle

    CoreProperties::CoreProperties(libxl::CoreProperties* coreProperties, Local<Value> book)
        : Wrapper<libxl::CoreProperties>(coreProperties), BookWrapper(book) {}

    Local<Object> CoreProperties::NewInstance(libxl::CoreProperties* libxlCoreProperties,
                                              Local<Value> book) {
        Nan::EscapableHandleScope scope;

        CoreProperties* coreProperties = new CoreProperties(libxlCoreProperties, book);

        Local<Object> that = util::CallStubConstructor(Nan::New(constructor)).As<Object>();

        coreProperties->Wrap(that);

        return scope.Escape(that);
    }

    // Wrappers

    NAN_METHOD(CoreProperties::Title) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        const char* title = that->GetWrapped()->title();
        if (!title) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(title).ToLocalChecked());
    }

    NAN_METHOD(CoreProperties::SetTitle) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(title, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setTitle(*title);

        info.GetReturnValue().Set(info.This());
    }

    // Init

    void CoreProperties::Initialize(Local<Object> exports) {
        using namespace libxl;

        Nan::HandleScope scope;

        Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(util::StubConstructor);
        t->SetClassName(Nan::New<String>("CoreProperties").ToLocalChecked());
        t->InstanceTemplate()->SetInternalFieldCount(1);

        BookWrapper::Initialize<CoreProperties>(t);

        Nan::SetPrototypeMethod(t, "title", Title);
        Nan::SetPrototypeMethod(t, "setTitle", SetTitle);

        t->ReadOnlyPrototype();
        constructor.Reset(Nan::GetFunction(t).ToLocalChecked());
    }
}  // namespace node_libxl